using Azure.Identity;
using Microsoft.Graph;
using TeamsGetMessages;

var settingReader = new SecretAppSettingReader();
var settings = settingReader.ReadSection<Settings>("Settings");


GraphServiceClient graphClient = CreateGraphClient();

var user = await graphClient.Me
    .Request()
    .GetAsync();

try
{
    var msg = await SendMessageToChannel(graphClient, settings.TeamId, settings.ChannelId, "Hi There");
}
catch (Exception ex)
{
    Console.WriteLine(ex.ToString());   
}


var subscription = new Subscription
{
    ChangeType = "created",
    NotificationUrl = settings.NotificationUrl,
    Resource = $"/teams/{settings.TeamId}/channels/{settings.ChannelId}/messages",
    ExpirationDateTime = new DateTimeOffset(DateTime.UtcNow.AddMinutes(59)),
    ClientState = "secretClientValue",
    LatestSupportedTlsVersion = "v1_2" //,
 //   IncludeResourceData = true
};

try
{
    var newSubscription = await graphClient.Subscriptions
                .Request()
                .AddAsync(subscription);
    Console.WriteLine($"New subscriptionID:  {newSubscription.Id.ToString()}");
}
catch(Exception ex)
{
    Console.WriteLine(ex.ToString());
}



static async Task<ChatMessage> SendMessageToChannel(GraphServiceClient graphClient, string teamId, string channelId, string message)
{
    var chatMessage = new ChatMessage
    {
        Body = new ItemBody
        {
            Content = message
        }
    };

    return await graphClient.Teams[teamId].Channels[channelId].Messages
        .Request()
        .AddAsync(chatMessage);

}

static GraphServiceClient CreateGraphClient()
{
    var settingReader = new SecretAppSettingReader();
    var settings = settingReader.ReadSection<Settings>("Settings");

    var scopes = new[] { "User.Read", "https://graph.microsoft.com/ChannelMessage.Send", "https://graph.microsoft.com/ChannelMessage.Read.All" };

    var options = new InteractiveBrowserCredentialOptions
    {
        TenantId = settings.TenantId,
        ClientId = settings.ClientId,
        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
        RedirectUri = new Uri("http://localhost")
    };

    var interactiveCredential = new InteractiveBrowserCredential(options);

    var graphClient = new GraphServiceClient(interactiveCredential, scopes);
    return graphClient;
}