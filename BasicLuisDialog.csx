using System;
using System.Threading.Tasks;

using Microsoft.Bot.Builder.Azure;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Luis;
using Microsoft.Bot.Builder.Luis.Models;

// For more information about this template visit http://aka.ms/azurebots-csharp-luis
[Serializable]
public class BasicLuisDialog : LuisDialog<object>
{

    public BasicLuisDialog() : base(new LuisService(new LuisModelAttribute(Utils.GetAppSetting("LuisAppId"), Utils.GetAppSetting("LuisAPIKey"))))
    {
    }


    const string SPAccessTokenKey = "SPAccessToken";
    const string SPSite = "https://nagarro.sharepoint.com/sites/teams/development";
    private static readonly Dictionary<string, string> PropertyMappings = new Dictionary<string, string>
        {
            { "TypeOfDocument", "botDocType" },
            { "Software", "BotTopic" }
        };

    [Serializable]
    public class PartialMessage
    {
        public string Text { set; get; }
    }
    private PartialMessage message;


    protected override async Task MessageReceived(IDialogContext context,
        IAwaitable<Microsoft.Bot.Connector.IMessageActivity> item)
    {
        var msg = await item;

        if (string.IsNullOrEmpty(context.UserData.Get<string>(SPAccessTokenKey)))
        {
            MicrosoftAppCredentials cred = new MicrosoftAppCredentials(
                ConfigurationManager.AppSettings["MicrosoftAppId"],
                ConfigurationManager.AppSettings["MicrosoftAppPassword"]);
            StateClient stateClient = new StateClient(cred);
            BotState botState = new BotState(stateClient);
            BotData botData = await botState.GetUserDataAsync(msg.ChannelId, msg.From.Id);
            context.UserData.SetValue<string>(SPAccessTokenKey, botData.GetProperty<string>(SPAccessTokenKey));
        }

        this.message = new PartialMessage { Text = msg.Text };
        await base.MessageReceived(context, item);
    }
    [LuisIntent("None")]
    public async Task NoneIntent(IDialogContext context, LuisResult result)
    {
        await context.PostAsync($"You have reached the none intent. You said: {result.Query}"); //
        context.Wait(MessageReceived);
    }

    // Go to https://luis.ai and create a new intent, then train/publish your luis app.
    // Finally replace "MyIntent" with the name of your newly created intent in the following handler
    [LuisIntent("FindDocuments")]
    public async Task MyIntent(IDialogContext context, LuisResult result)
    {
        var reply = context.MakeMessage();
        try
        {
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            reply.Attachments = new List<Microsoft.Bot.Connector.Attachment>();
            StringBuilder query = new StringBuilder();
            bool QueryTransformed = false;
            if (result.Entities.Count > 0)
            {
                QueryTransformed = true;
                foreach (var entity in result.Entities)
                {
                    if (PropertyMappings.ContainsKey(entity.Type))
                    {
                        query.AppendFormat("{0}:'{1}' ", PropertyMappings[entity.Type], entity.Entity);
                    }
                }
            }
            else
            {
                //should replace all special chars
                query.Append(this.message.Text.Replace("?", ""));
            }

            using (ClientContext ctx = new ClientContext(SPSite))
            {
                ctx.AuthenticationMode = ClientAuthenticationMode.Anonymous;
                ctx.ExecutingWebRequest +=
                    delegate (object oSender, WebRequestEventArgs webRequestEventArgs)
                    {
                        webRequestEventArgs.WebRequestExecutor.RequestHeaders["Authorization"] =
                            "Bearer " + context.UserData.Get<string>("SPAccessToken");
                    };
                KeywordQuery kq = new KeywordQuery(ctx);
                kq.QueryText = string.Concat(query.ToString(), " IsDocument:1");
                kq.RowLimit = 5;
                SearchExecutor se = new SearchExecutor(ctx);
                ClientResult<ResultTableCollection> results = se.ExecuteQuery(kq);
                ctx.ExecuteQuery();

                if (results.Value != null && results.Value.Count > 0 && results.Value[0].RowCount > 0)
                {
                    reply.Text += (QueryTransformed == true) ? "I found some interesting reading for you!" : "I found some potential interesting reading for you!";
                    BuildReply(results, reply);
                }
                else
                {
                    if (QueryTransformed)
                    {
                        //fallback with the original message
                        kq.QueryText = string.Concat(this.message.Text.Replace("?", ""), " IsDocument:1");
                        kq.RowLimit = 3;
                        se = new SearchExecutor(ctx);
                        results = se.ExecuteQuery(kq);
                        ctx.ExecuteQuery();
                        if (results.Value != null && results.Value.Count > 0 && results.Value[0].RowCount > 0)
                        {
                            reply.Text += "I found some potential interesting reading for you!";
                            BuildReply(results, reply);
                        }
                        else
                            reply.Text += "I could not find any interesting document!";
                    }
                    else
                        reply.Text += "I could not find any interesting document!";

                }

            }

        }
        catch (Exception ex)
        {
            reply.Text = ex.Message;
        }
        await context.PostAsync(reply);
        context.Wait(MessageReceived);
    }
    void BuildReply(ClientResult<ResultTableCollection> results, IMessageActivity reply)
    {
        foreach (var row in results.Value[0].ResultRows)
        {
            List<CardAction> cardButtons = new List<CardAction>();
            List<CardImage> cardImages = new List<CardImage>();
            string ct = string.Empty;
            string icon = string.Empty;
            switch (row["FileExtension"].ToString())
            {
                case "docx":
                    ct = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                    icon = "https://cdn2.iconfinder.com/data/icons/metro-ui-icon-set/128/Word_15.png";
                    break;
                case "xlsx":
                    ct = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    icon = "https://cdn2.iconfinder.com/data/icons/metro-ui-icon-set/128/Excel_15.png";
                    break;
                case "pptx":
                    ct = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
                    icon = "https://cdn2.iconfinder.com/data/icons/metro-ui-icon-set/128/PowerPoint_15.png";
                    break;
                case "pdf":
                    ct = "application/pdf";
                    icon = "https://cdn4.iconfinder.com/data/icons/CS5/256/ACP_PDF%202_file_document.png";
                    break;

            }
            cardButtons.Add(new CardAction
            {
                Title = "Open",
                Value = (row["ServerRedirectedURL"] != null) ? row["ServerRedirectedURL"].ToString() : row["Path"].ToString(),
                Type = ActionTypes.OpenUrl
            });
            cardImages.Add(new CardImage(url: icon));
            ThumbnailCard tc = new ThumbnailCard();
            tc.Title = (row["Title"] != null) ? row["Title"].ToString() : "Untitled";
            tc.Text = (row["Description"] != null) ? row["Description"].ToString() : string.Empty;
            tc.Images = cardImages;
            tc.Buttons = cardButtons;
            reply.Attachments.Add(tc.ToAttachment());
        }
    }

}