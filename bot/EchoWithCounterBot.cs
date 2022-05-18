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
using Azure;
using Azure.AI.TextAnalytics;
using Azure.Search.Documents;
using Azure.Search.Documents.Models;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Configuration;
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

        // stopwords (words that are never used in a search query even if present in a user's question)
        private static string stopwords = "what who whom which what when how was does this that the mean means meaning cryptonym crypt code name word codename codeword";
        private static HashSet<string> stopset;

        //  set that remembers who the bot has greeted (add default-user to avoid double-greeting)
        private static HashSet<string> greeted;

        // known cryptonyms (code names) read from JSON file
        private static Dictionary<string, string> cryptonyms;

        // client interfaces for Cognitive Services APIs
        private static TextAnalyticsClient textAnalyticsClient;
        private static SearchClient searchClient;

        // search URL for your main JFK Files site instance
        private static string searchUrl;

        private Activity activity;
        private ITurnContext context;
        private Guid turnid;

        private ILogger logger;

        /// <summary>
        /// Initializes the static members of the EchoWithCounterBot class
        /// </summary>
        static EchoWithCounterBot()
        {
            var config = Startup.Configuration;

            var textAnalyticsKey = config.GetSection("textAnalyticsKey")?.Value;
            var textAnalyticsEndpoint = config.GetSection("textAnalyticsEndpoint")?.Value;

            textAnalyticsClient = new TextAnalyticsClient(new Uri(textAnalyticsEndpoint), new AzureKeyCredential(textAnalyticsKey));

            var searchEndpoint = config.GetSection("searchUrl")?.Value;
            var searchIndex = config.GetSection("searchIndex")?.Value;
            var searchKey = config.GetSection("searchKey")?.Value;

            searchClient = new SearchClient(new Uri(searchEndpoint), searchIndex, new AzureKeyCredential(searchKey));

            stopset = new HashSet<string>(stopwords.Split());
            greeted = new HashSet<string>() { "default-user" };

            cryptonyms = JsonConvert.DeserializeObject<Dictionary<string, string>>
                (File.ReadAllText("cia-cryptonyms.json"));

            searchUrl = config.GetSection("searchUrl")?.Value;
        }

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
            // member variables -- valid for this turn only
            context = turnContext;
            activity = context.Activity;

            var type = activity.Type;

            logger.LogTrace($"HOOVERBOT {turnid} received {type} from {activity.From.Id}");

            // Send initial greeting on ConversationUpdate activity
            if (type == ActivityTypes.ConversationUpdate)
            {
                SendInitialGreeting();
                return;
            }

            // Process message from user on Message activity
            if (type == ActivityTypes.Message)
            {
                var question = activity.Text;
                logger.LogTrace($"HOOVERBOT {turnid} received message: {question}.");

                // handle greetings, gratitude, welcome, etc.
                if (HandleSocialNiceties(question.ToLower())) return;

                // create HashSet for search terms
                var terms = new HashSet<string>();

                // handle cryptonyms, sending a definition and adding them to our search terms
                terms.UnionWith(ExtractAndDefineCryptonyms(question));
                var crypt_found = terms.Count > 0;

                // extract search keywords from question and add them to search terms
                await ExtractKeywords(question, terms);

                // Build the search term
                var query = string.Join(" ", terms).Trim();

                // if we failed to build a search, punt back to the user's original query
                if (query == string.Empty)
                {
                    query = question;
                }

                // Perform Azure search of the JFK Files documents
                var results = await PerformSearch(query);

                logger.LogTrace($"HOOVERBOT {turnid} received {results.TotalCount} search results");
                SendTypingIndicator();  // to cover building and sending the response

                // Build a reply with the search results
                var reply = BuildResultsReply(results, query, crypt_found);

                // send the reply if we have search results OR if we didn't find a cryptonym
                // if we found a cryptonym but no search results, no need to send "no results"
                // (the cryptonym hit IS a result, and the user will find "no results" puzzling)
                if (reply.Attachments.Count > 1 || !crypt_found)
                {
                    reply.Properties["cache-speech"] = true;    // allows client to retain speech audio
                    context.SendActivityAsync(reply);
                    logger.LogTrace($"HOOVERBOT {turnid} sent search response");
                }
            }

        }

        //
        // METHODS for processing the user's question, performing a search, and sending a reply
        //

        // Respond to greetings and social niceties.
        // return value indicates whether message was handled (so we can bail out early)
        private bool HandleSocialNiceties(string question)
        {
            if (question.Contains("welcome"))
            {
                logger.LogTrace($"HOOVERBOT {turnid} responding to welcome.");
                SendCacheableSpeechReply(Thanks);
                return true;
            }

            if (question.Contains("hello"))
            {
                logger.LogTrace($"HOOVERBOT {turnid} responding to hello.");
                SendCacheableSpeechReply(Hello);
                return true;
            }

            if (question.Contains("thank"))
            {
                logger.LogTrace($"HOOVERBOT {turnid} responding to thanks.");
                SendCacheableSpeechReply(YoureWelcome);
                return true;
            }

            return false;
        }

        // Extract cryptonyms fom the user's question.
        // These are added to the search query and a definition is sent for each.
        private HashSet<string> ExtractAndDefineCryptonyms(string question)
        {
            var terms = new HashSet<string>();

            // look for cryptomyms and send the definition of any found
            Regex words = new Regex(@"\b(\w+)\b", RegexOptions.Compiled);
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
                }
            }

            logger.LogTrace($"HOOVERBOT {turnid} found cryptonyms {string.Join(" ", terms)}.");
            return terms;
        }

        // Extract search keywords from the user's question using the Text Analytics service
        // Both Key Phrases and Entities are extracted
        private async Task<HashSet<string>> ExtractKeywords(string question, HashSet<string> terms)
        {

            // do the Text Analytics requests asynchronously so both can proceed at the same time
            logger.LogTrace($"HOOVERBOT {turnid} making Text Analytics requests.");
            var keyphrases = await textAnalyticsClient.ExtractKeyPhrasesAsync(question);
            var entities = await textAnalyticsClient.RecognizeEntitiesAsync(question);
            var keyphrases = await t_keyphrases;
            var entities = await t_entities;
            logger.LogTrace($"HOOVERBOT {turnid} received Text Analytics responses.");

            // helper function to process individual words found by Text Analytics, ignoring some
            // * possesive words have their trailing 's or ' removed (Kenneydy's -> Kennedy)
            // * words ending with "-ment" or "-ion" are ignored (these are often picked up by Text Analytics but have little search value)
            // * words whose uppercase variant is already in the search terms are ignored (these are typically cryptonyms)
            // * words of only 1 or 2 letters long are ignored
            // * stopwords are ignored
            void addTerm(string word)
            {
                if (word.EndsWith("'s"))
                {
                    word = word.TrimEnd('s');
                }
                if (word.EndsWith("'"))
                {
                    word = word.TrimEnd('\'');
                }
                if (word.Length > 2 && !(word.EndsWith("ment") || word.EndsWith("ion") || terms.Contains(word.ToUpper()) || stopset.Contains(word.ToLower())))
                {
                    terms.Add(word);
                }
            }

            foreach (var phrase in keyphrases.Value)
            {
                logger.LogTrace($"HOOVERBOT {turnid} processing Key Phrase: {phrase}");
                foreach (var word in phrase.Split())
                {
                    addTerm(word);
                } 
            }

            // we don't directly use the names of recognized entities. instead, we use the term from the user's query that was recognized as an entity.
            // that is, if they entered Oswald, the recognized entiity is Lee Harvey Oswald, but we add the user's term, JFK, to the query.
            // in other words, certain words beinga recognized as entities simply tells us they're important to search for, but we still search for what the user entered.
            foreach (var entity in entities.Value)
            {
                logger.LogTrace($"HOOVERBOT {turnid} processing Entity SubCategory: {entity.SubCategory}");
                foreach (var word in entity.Text.Split())
                {
                    addTerm(word);
                }
            }

            return terms;
        }

        // Perform a search (send a typing indicator every 2 seconds while waiting for results)
        private async Task<SearchResults<SearchDocument>> PerformSearch(string query)
        {
            // initiate the search
            SearchOptions options = new SearchOptions()
            {
                Size = MaxResults,
                IncludeTotalCount = true
            };   // get top n results
            var search = searchClient.SearchAsync<SearchDocument>(query, options);
            logger.LogTrace($"HOOVERBOT {turnid} searching JFK Files for derived query: {query}");

            // send typing indicator every 2 sec while we wait for search to complete
            do
            {
                SendTypingIndicator();
            }
            while (!search.Wait(2000));
            
            return search.Result.Value;
        }

        // Build a custom carousel reply with a card for each search result
        // The card displays a thumbnail image of the document and a brief text excerpt
        private Activity BuildResultsReply(SearchResults<SearchDocument> results, string query, bool crypt_found)
        {
            // create a reply
            var reply = activity.CreateReply();
            reply.Attachments = new List<Attachment>();

            foreach (var result in results.GetResults())
            {
                // get enrichment data from search result (and parse the JSON data)
                var enriched = JObject.Parse((string)result.Document["enriched"]);

                // all results should have thumbnail images, but just in case, look before leaping
                if (enriched.TryGetValue("/document/normalized_images/*/imageStoreUri", out var images))
                {
                    // Get URI of thumbnail of first content page.
                    // If the document has multiple pages, first page is an identification form
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

            DescribeResultsReply(reply, query, crypt_found);
            AddDigDeeperButton(reply, query);
            logger.LogTrace($"HOOVERBOT {turnid} added text={reply.Text} and speak={reply.Speak}");
            return reply;
        }

        // Add a text element describing the results from the search, if any.
        private void DescribeResultsReply(Activity reply, string query, bool crypt_found)
        {

            // Add text describing results, if any
            if (reply.Attachments.Count == 0)
            {
                reply.Text = $"Sorry, I didn't find any documents matching **{query}**.";
                reply.Speak = "Sorry, I didn't find any documents about that.";
            }
            else
            {
                var documents = reply.Attachments.Count > 1 ? "some documents" : "a document";
                reply.Text = $"I found {documents} about **{query}** you may be interested in.";
                reply.Speak = $"I found {documents} you may be interested in.";
                if (crypt_found)
                {
                    reply.Speak = "Also, " + reply.Speak;
                }

                reply.Properties["cache-speech"] = true;
                reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            }
        }

        // Add a "Dig Deeper" button to the search results reply
        // Users can click this button to search on the JFK Files Web site
        private void AddDigDeeperButton(Activity reply, string query)
        {
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

        // Send initial greeting
        // Each user in the chat (including the bot) is added via a ConversationUpdate message.
        // Greet only once per conversation (appropriate for Web Chat, might not be good for other channels)
        private async void SendInitialGreeting()
        {
            if (activity.MembersAdded != null)
            {
                foreach (var member in activity.MembersAdded)
                {
                    if (member.Id != activity.Recipient.Id && !greeted.Contains(activity.Conversation.Id))
                    {
                        SendTextOnlyReply(Greeting);
                        greeted.Add(activity.Conversation.Id);
                        logger.LogTrace($"HOOVERBOT {turnid} greeted user: {member.Name} {member.Id}");
                        break;
                    }
                }
            }
        }   

        //
        // CONVENIENCE METHODS for sending replies
        //

        // Send a text-only reply (no speech)
        private async void SendTextOnlyReply(string text)
        {
            _SendReply(text, null, false);
            logger.LogTrace($"HOOVERBOT {turnid} sent text-only reply: {text.Substring(0, Math.Min(50, text.Length))}");
        }

        // Send a reply with a speak attribute so it will be spoken by the client
        // Optionally specify text to be spoken if it is different from the text to be displayed
        // Optionally set caching attribute so boilerplate speech responses may be cached client-side
        private async void SendSpeechReply(string text, string speech=null)
        {
            _SendReply(text, speech ?? text, false);
            logger.LogTrace($"HOOVERBOT {turnid} sent speech reply: {text.Substring(0, Math.Min(50, text.Length))}");
        }

        private async void SendCacheableSpeechReply(string text, string speech=null)
        {
            _SendReply(text, speech ?? text, true);
            logger.LogTrace($"HOOVERBOT {turnid} sent cacheable speech reply: {text.Substring(0, Math.Min(50, text.Length))}");
        }

        // Send a typing indicator while other work (e.g. database search) is in progress.
        // A single typing indicator message keeps the "..." animation visible for up to three seconds
        // or until next message is received by the client.
        private async void SendTypingIndicator()
        {
            var typing = activity.CreateReply();
            typing.Type = ActivityTypes.Typing;
            context.SendActivityAsync(typing);
            logger.LogTrace($"HOOVERBOT {turnid} sent typing indicator");
        }

        // Helper method called by convenience methods. Call one of those, not this one.
        private async void _SendReply(string text, string speech, bool cacheable)
        {
            var reply = activity.CreateReply();
            reply.Text = text;
            if (speech != null && speech.Length > 0) reply.Speak = speech;
            if (cacheable) reply.Properties["cache-speech"] = true;
            context.SendActivityAsync(reply);
        }

        //
        // RESULT CARD layout class
        //

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
