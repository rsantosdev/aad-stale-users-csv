using System.Globalization;
using Azure.Identity;
using CsvHelper;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace AzureAdStaleUsers;

internal class Program
{
    static async Task Main(string[] args)
    {
        var config = new ConfigurationBuilder()
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
            .Build();

        var scopes = new[] { "User.Read.All", "AuditLog.Read.All" };

        // using Azure.Identity;
        var options = new DeviceCodeCredentialOptions
        {
            AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
            ClientId = config["ClientId"],
            TenantId = config["TenantId"],
            DeviceCodeCallback = (code, cancellation) =>
            {
                Console.WriteLine(code.Message);
                return Task.FromResult(0);
            },
        };

        // https://learn.microsoft.com/dotnet/api/azure.identity.devicecodecredential
        var deviceCodeCredential = new DeviceCodeCredential(options);

        var graphClient = new GraphServiceClient(deviceCodeCredential, scopes);

        int.TryParse(config["StaleDays"], out var staleDays);
        var staleDate = DateTime.Today.AddDays(staleDays * -1).ToString("yyyy-MM-dd");
        var result = await graphClient.Users.GetAsync((requestConfiguration) =>
        {
            requestConfiguration.QueryParameters.Select = ["userPrincipalName", "displayName", "mail", "signInActivity"];
            requestConfiguration.QueryParameters.Filter = $"signInActivity/lastSignInDateTime le {staleDate}T00:00:00Z";
        });

        if (result == null)
        {
            Console.WriteLine("No users found.");
            return;
        }

        var finalList = new List<StaleUser>();
        var pageIterator = PageIterator<User, UserCollectionResponse>
            .CreatePageIterator(
                graphClient,
                result,
                item =>
                {
                    finalList.Add(new StaleUser
                    {
                        Id = item.Id,
                        DisplayName = item.DisplayName,
                        Mail = item.Mail ?? item.UserPrincipalName,
                        LastSignInDateTime = item.SignInActivity?.LastSignInDateTime?.ToString("yyyy-MM-dd"),
                    });

                    return true;
                });

        await pageIterator.IterateAsync();

        await using (var writer = new StreamWriter($"{DateTime.Now:yyyy-dd-M--HH-mm-ss}.csv"))
        await using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
        {
            await csv.WriteRecordsAsync(finalList);
        }

        Console.WriteLine("DONE!");
    }
}

internal class StaleUser
{
    public string? Id { get; set; }
    public string? DisplayName { get; set; }
    public string? Mail { get; set; }
    public string? LastSignInDateTime { get; set; }
}