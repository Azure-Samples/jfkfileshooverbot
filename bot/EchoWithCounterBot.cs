// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using AdaptiveCards;
using Microsoft.Azure.CognitiveServices.Language.TextAnalytics;
using Microsoft.Azure.CognitiveServices.Language.TextAnalytics.Models;
using Microsoft.Azure.Search;
using Microsoft.Azure.Search.Models;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Logging;
using Microsoft.Rest;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

#pragma warning disable 4014
#pragma warning disable 1998

namespace Microsoft.BotBuilderSamples
{
    /// <summary>
    /// Represents a bot that processes incoming activities.
    /// For each user interaction, an instance of this class is created and the OnTurnAsync method is called.
    /// This is a Transient lifetime service.  Transient lifetime services are created
    /// each time they're requested. For each Activity received, a new instance of this
    /// class is created. Objects that are expensive to construct, or have a lifetime
    /// beyond the single turn, should be carefully managed.
    /// For example, the <see cref="MemoryStorage"/> object and associated
    /// <see cref="IStatePropertyAccessor{T}"/> object are created with a singleton lifetime.
    /// </summary>
    /// <seealso cref="https://docs.microsoft.com/en-us/aspnet/core/fundamentals/dependency-injection?view=aspnetcore-2.1"/>
    public class EchoWithCounterBot : IBot
    {
        private static readonly int MaxResults = 10;
        private static readonly string Greeting = "Welcome, investigator! I'm FBI director J. Edgar Hoover. What can I help you find?";
        private static readonly string Thanks = "Thank you, it's my pleasure to be here. What can I do for you?";
        private static readonly string Hello = "Hello! Good to meet you. What are you interested in today?";
        private static readonly string YoureWelcome = "You're most welcome. How can I assist you?";

        private static readonly object InitLock = new object();

        private static SearchIndexClient searchClient = null;
        private static Dictionary<string, string> cryptonyms = null;
        private static string searchUrl = null;
        private static HashSet<string> greeted;

        private static ITextAnalyticsClient textAnalyticsClient;

        private Activity activity;
        private ITurnContext context;
        private Guid turnid;

        private ILogger logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="EchoWithCounterBot"/> class.
        /// </summary>
        /// <param name="accessors">A class containing <see cref="IStatePropertyAccessor{T}"/> used to manage state.</param>
        /// <param name="loggerFactory">A <see cref="ILoggerFactory"/> that is hooked to the Azure App Service provider.</param>
        /// <seealso cref="https://docs.microsoft.com/en-us/aspnet/core/fundamentals/logging/?view=aspnetcore-2.1#windows-eventlog-provider"/>
        public EchoWithCounterBot(EchoBotAccessors accessors, ILoggerFactory loggerFactory)
        {
            turnid = System.Guid.NewGuid();

            if (loggerFactory == null)
            {
                throw new System.ArgumentNullException(nameof(loggerFactory));
            }

            logger = loggerFactory.CreateLogger<EchoWithCounterBot>();
            logger.LogTrace($"HOOVERBOT {turnid} turn start.");

            // avoid multiple initialization of static fields
            // it doesn't hurt anything in this case, but it's bad practice
            lock (InitLock)
            {
                if (greeted == null)
                {
                    var config = Startup.Configuration;

                    var searchname = config.GetSection("searchName")?.Value;
                    var searchkey = config.GetSection("searchKey")?.Value;
                    var searchindex = config.GetSection("searchIndex")?.Value;

                    // establish search service connection
                    searchClient = new SearchIndexClient(searchname, searchindex, new SearchCredentials(searchkey));

                    // read known cryptonyms (code names) from JSON file
                    cryptonyms = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText("cia-cryptonyms.json"));

                    // get search URL for your main JFK Files site instance
                    searchUrl = config.GetSection("searchUrl")?.Value;

                    // create Text Analytics client
                    var textAnalyticsKey = config.GetSection("textAnalyticsKey")?.Value;
                    var textAnalyticsEndpoint = config.GetSection("textAnalyticsEndpoint")?.Value;
                    textAnalyticsClient = new TextAnalyticsClient(new ApiKeyServiceClientCredentials(textAnalyticsKey))
                    {
                        Endpoint = textAnalyticsEndpoint,
                    };

                    // create set that remembers who the bot has greeted (add default-user to avoid double-greeting)
                    greeted = new HashSet<string>() { "default-user" };
                }
            }
        }

        /// <summary>
        /// Every conversation turn for our Bot will call this method.
        /// There are no dialogs used, since it's "single turn" processing, meaning a single
        /// request and response.
        /// </summary>
        /// <param name="turnContext">A <see cref="ITurnContext"/> containing all the data needed
        /// for processing this conversation turn. </param>
        /// <param name="cancellationToken">(Optional) A <see cref="CancellationToken"/> that can be used by other objects
        /// or threads to receive notice of cancellation.</param>
        /// <returns>A <see cref="Task"/> that represents the work queued to execute.</returns>
        /// <seealso cref="BotStateSet"/>
        /// <seealso cref="ConversationState"/>
        /// <seealso cref="IMiddleware"/>
        public async Task OnTurnAsync(ITurnContext turnContext, CancellationToken cancellationToken)
        {
            context = turnContext;
            activity = context.Activity;
            var type = activity.Type;

            // add bot's ID to greeted set so we don't greet it (no harm in doing it eeach time)
            greeted.Add(activity.Recipient.Id);

            if (type == ActivityTypes.Message)
            {
                // respond to greetings and social niceties
                var question = activity.Text.ToLower();
                logger.LogTrace($"HOOVERBOT {turnid} received message: {question}.");

                if (question.Contains("welcome"))
                {
                    logger.LogTrace($"HOOVERBOT {turnid} responding to welcome.");
                    SendCacheableSpeechReply(Thanks);
                    return;
                }

                if (question.Contains("hello"))
                {
                    logger.LogTrace($"HOOVERBOT {turnid} responding to hello.");
                    SendCacheableSpeechReply(Hello);
                    return;
                }

                if (question.Contains("thank"))
                {
                    logger.LogTrace($"HOOVERBOT {turnid} responding to thanks.");
                    SendCacheableSpeechReply(YoureWelcome);
                    return;
                }

                question = activity.Text;   // back to our original value with case as user entered it

                // HashSet will hold the search terms for the Azure Search query
                // We use a hashset since we want to include each term only once
                var terms = new HashSet<string>();

                // look for cryptomyms and send the definition of any found
                Regex words = new Regex(@"\b(\w+)\b", RegexOptions.Compiled);
                var crypt_found = false;
                foreach (Match match in words.Matches(question))
                {
                    var word = match.Groups[0].Value;
                    var upperword = word.ToUpper();
                    if (cryptonyms.ContainsKey(upperword))
                    {
                        // uppercase cryptonym in the question
                        question = new Regex("\\b" + word + "\\b").Replace(question, upperword);
                        terms.Add(upperword);       // add them to our Azure Search terms
                        SendSpeechReply($"**{upperword}**: {cryptonyms[upperword]}", cryptonyms[upperword]);
                        crypt_found = true;
                    }
                }

                logger.LogTrace($"HOOVERBOT {turnid} found cryptonyms {string.Join(" ", terms)}.");

                // use Text Analytics to get key phrases and entities from the question, which will become the search query
                // we will use the same Text Analytics query for both key words and entities
                var doc = new MultiLanguageBatchInput(
                    new List<MultiLanguageInput>()
                    {
                        new MultiLanguageInput("en", "0", question),
                    });

                // do the Text Analytics requests asynchronously so both can proceed at the same time
                logger.LogTrace($"HOOVERBOT {turnid} making Text Analytics requests.");
                var t_keyphrases = textAnalyticsClient.KeyPhrasesAsync(doc);
                var t_entities = textAnalyticsClient.EntitiesAsync(doc);
                var keyphrases = await t_keyphrases;
                var entities = await t_entities;
                logger.LogTrace($"HOOVERBOT {turnid} received Text Analytics responses.");

                // build up the query string using the results from the Text Analytics queries

                // PROCESS KEY PHRASES
                // we'll tweak the returned key phrases as they are not exactly what we want for a search
                // any term of form "-ing of X" (e.g. "meaning of GPIDEAL") will be skipped since this is probably a cryptonym we've already handled
                // words ending in 's or ' (e.g. "Kennedy's") will have their possessive removed
                // words ending in -ment or -ion will be omitted entirely; these are often recognized as parts of keywords but usually are not useful search terms
                foreach (var phrase in keyphrases.Documents[0].KeyPhrases)
                {
                    logger.LogTrace($"HOOVERBOT {turnid} processing Key Phrase: {phrase}");
                    if (!phrase.Contains("ing of "))
                    {
                        foreach (var word in phrase.Split())
                        {
                            if (word.EndsWith("'s"))
                            {
                                terms.Add(word.TrimEnd('s').TrimEnd('\''));
                            }
                            else if (word.EndsWith("'"))
                            {
                                terms.Add(word.TrimEnd('\''));
                            }
                            else if (!word.EndsWith("ment") && !word.EndsWith("ion"))
                            {
                                terms.Add(word);
                            }
                        }
                    }
                }

                // PROCESS ENTITIES
                // we don't directly use the names of recognized entities. instead, we use the term from the user's query that was recognized as an entity.
                // that is, if they entered Oswald, the recognized entiity is Lee Harvey Oswald, but we add the user's term, JFK, to the query.
                // in other words, certain words beinga recognized as entities simply tells us they're important to search for, but we still search for what the user entered.
                foreach (var entity in entities.Documents[0].Entities)
                {
                    logger.LogTrace($"HOOVERBOT {turnid} processing Entity: {entity.Name}");
                    foreach (var match in entity.Matches)
                    {
                        logger.LogTrace($"HOOVERBOT {turnid} Entity {entity.Name} derived from: {match.Text}");
                        terms.UnionWith(match.Text.Split());
                    }
                }

                var query = string.Join(" ", terms).Trim();
                var displayquery = query;

                // if we failed to build a search, punt back to the user's original query
                if (query == string.Empty)
                {
                    displayquery = query = question;
                }

                // initiate the search
                var parameters = new SearchParameters() { Top = MaxResults };   // get top n results
                var search = searchClient.Documents.SearchAsync(query, parameters);
                logger.LogTrace($"HOOVERBOT {turnid} searching JFK Files for derived query: {query}");

                // send typing indicator every 2 sec while we wait for search to complete
                do
                {
                    SendTypingIndicator();
                }
                while (!search.Wait(2000));

                var results = search.Result.Results;
                SendTypingIndicator();  // to cover building and sending the response
                logger.LogTrace($"HOOVERBOT {turnid} received {results.Count} search results");

                // create a reply
                var reply = activity.CreateReply();
                reply.Attachments = new List<Attachment>();

                foreach (var result in results)
                {
                    // get enrichment data from search result (and parse the JSON data)
                    var enriched = JObject.Parse((string)result.Document["enriched"]);

                    // all results should have thumbnail images, but just in case, look before leaping
                    if (enriched.TryGetValue("/document/normalized_images/*/imageStoreUri", out var images))
                    {
                        // get URI of thumbnail of first content page
                        // if the document has multiple pages, first page is an identification form
                        // and so the second page is the first page of interest; use its thumbnail
                        var thumbs = new List<string>(images.Values<string>());
                        var picurl = thumbs[thumbs.Count > 1 ? 1 : 0];

                        // get valid URL of original document (combine path and token)
                        var document = enriched["/document"];
                        var filename = document["metadata_storage_path"].Value<string>();
                        var token = document["metadata_storage_sas_token"].Value<string>();
                        var docurl = $"{filename}?{token}";

                        logger.LogTrace($"HOOVERBOT {turnid} added search result to response: {docurl}");

                        // Get the text from the document. This includes OCR'd printed and
                        // handwritten text, people recognized in photos, and more.
                        // As with the image, try to get the second page's text if it's multi-page
                        var text = enriched["/document/finalText"].Value<string>();
                        if (thumbs.Count > 1)
                        {
                            var sep = "[image: image1.tif]";
                            var page2 = text.IndexOf(sep);
                            text = page2 > 0 ? text.Substring(page2 + sep.Length) : text;
                        }

                        // create card for this search result and attach it to the reply
                        var card = new ResultCard(picurl, text, docurl);
                        reply.Attachments.Add(card.ToAttachment());
                    }
                }

                // Add text describing results, if any
                if (reply.Attachments.Count == 0)
                {
                    reply.Text = $"Sorry, I didn't find any documents matching __{displayquery}__.";
                    reply.Speak = "Sorry, I didn't find any documents about that.";
                }
                else
                {
                    var documents = reply.Attachments.Count > 1 ? "some documents" : "a document";
                    reply.Text = $"I found {documents} about **{displayquery}** you may be interested in.";
                    reply.Speak = $"I found {documents} you may be interested in.";
                    if (crypt_found)
                    {
                        reply.Speak = "Also, " + reply.Speak;
                    }

                    reply.Properties["cache-speech"] = true;
                    reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;

                    // add "Dig Deeper" button
                    var searchlink = searchUrl + System.Uri.EscapeDataString(query);
                    reply.SuggestedActions = new SuggestedActions()
                    {
                        Actions = new List<CardAction>()
                        {
                            new CardAction()
                            {
                                Title = "Dig Deeper",
                                Type = ActionTypes.OpenUrl,
                                Value = searchlink,
                            },
                        },
                    };
                    logger.LogTrace($"HOOVERBOT {turnid} added Dig Deeper {searchlink}");
                }

                // send the reply if we have search results OR if we didn't find a cryptonym
                // if we found a cryptonym but no search results, no need to send "no results"
                // (the cryptonym hit IS a result, and the user will find "no results" puzzling)
                if (reply.Attachments.Count > 1 || !crypt_found)
                {
                    reply.Properties["cache-speech"] = true;
                    context.SendActivityAsync(reply);
                    logger.LogTrace($"HOOVERBOT {turnid} sent search response");
                }
            }

            // Send initial greeting
            // Each user in the chat (including the bot) is added via a ConversationUpdate message
            // Check each user to make sure it's not the bot before greeting, and only greet each user once
            else if (type == ActivityTypes.ConversationUpdate)
            {
                if (activity.MembersAdded != null)
                {
                    foreach (var member in activity.MembersAdded)
                    {
                        if (!greeted.Contains(member.Id))
                        {
                            SendTextOnlyReply(Greeting);
                            greeted.Add(member.Id);
                            logger.LogTrace($"HOOVERBOT {turnid} greeted user: {member.Name} {member.Id}");
                            break;
                        }
                    }
                }
            }
        }

        // Send a text-only reply (no speech)
        private void SendTextOnlyReply(string text, string speech = null)
        {
            var reply = activity.CreateReply();
            reply.Text = text;
            context.SendActivityAsync(reply);
            logger.LogTrace($"HOOVERBOT {turnid} sent text-only reply: {text.Substring(0, Math.Min(50, text.Length))}");
        }

        // Send a reply with a speak attribute so it will be spoken by the client
        private void SendSpeechReply(string text, string speech = null)
        {
            var reply = activity.CreateReply();
            reply.Text = text;
            reply.Speak = speech == null ? text : speech;
            context.SendActivityAsync(reply);
            logger.LogTrace($"HOOVERBOT {turnid} sent speech reply: {text.Substring(0, Math.Min(50, text.Length))}");
        }

        // Send a reply with a speak attribute so it will be spoken by the client
        // Also include a value attribute indicating that the speech may be cached
        // Common boilerplate replies are good candidates for caching
        private void SendCacheableSpeechReply(string text, string speech = null)
        {
            var reply = activity.CreateReply();
            reply.Text = text;
            reply.Properties["cache-speech"] = true;
            reply.Speak = speech == null ? text : speech;
            context.SendActivityAsync(reply);
            logger.LogTrace($"HOOVERBOT {turnid} sent cacheable speech reply: {text.Substring(0, Math.Min(50, text.Length))}");
        }

        // Send a typing indicator while other work (e.g. database search) is in progress.
        // A single typing indicator message keeps the "..." animation visible for up to three seconds
        // or until next message is received by the client.
        private void SendTypingIndicator()
        {
            var typing = activity.CreateReply();
            typing.Type = ActivityTypes.Typing;
            context.SendActivityAsync(typing);
            logger.LogTrace($"HOOVERBOT {turnid} sent typing indicator");
        }

        // Result card layout using a simple adaptive layout
        private class ResultCard : AdaptiveCard
        {
            private int maxlen = 200;

            public ResultCard(string image_url, string text, string action_url)
            {
                // some documents have megabytes of text; avoid sending it all back to client
                text = text.Length < maxlen ? text : text.Substring(0, maxlen - 1) + "…";

                // add spot for thumbnail image
                this.Body.Add(new Image()
                {
                    Url = image_url,
                    Size = ImageSize.Large,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    AltText = text,
                    SelectAction = new OpenUrlAction() { Url = action_url, },
                });

                // add spot for text from document
                this.Body.Add(new TextBlock()
                {
                    Text = text,
                    MaxLines = 5,   // doesn't seem to work
                    Separation = SeparationStyle.Strong,
                });
            }

            public Attachment ToAttachment()
            {
                return new Attachment()
                {
                    ContentType = AdaptiveCard.ContentType,
                    Content = this,
                };
            }
        }

        // HTTP client credentials class containing an API key (used with Text Analytics client)
        private class ApiKeyServiceClientCredentials : ServiceClientCredentials
        {
            private string subscriptionKey;

            public ApiKeyServiceClientCredentials(string key)
                : base()
            {
                subscriptionKey = key;
            }

            public override Task ProcessHttpRequestAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                request.Headers.Add("Ocp-Apim-Subscription-Key", subscriptionKey);
                return base.ProcessHttpRequestAsync(request, cancellationToken);
            }
        }
    }
}
