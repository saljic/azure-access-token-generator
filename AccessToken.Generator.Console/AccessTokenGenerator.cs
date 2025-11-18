using Azure.Core;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using OneOf;
using Spectre.Console;
using TextCopy;

namespace AccessToken.Generator.Console;

public sealed class AccessTokenGenerator(GraphServiceClient graphClient, TokenCredential tokenCredential)
{
    private Application[] _cachedApplications = [];
    
    public async Task GenerateAsync(CancellationToken token = default)
    {
        AnsiConsole.Clear();
        
        bool isRunning = true;

        while (isRunning)
        {
            var applications = await GetApplicationsAsync(token);
            var applicationsWithAppRole =
                applications.Where(x => x.AppRoles.Any() && x.IdentifierUris.Any() && x.DisplayName is not null);

            var filteredApplications = FilterWithFuzzySearch(
                applicationsWithAppRole.Select(x => x.DisplayName).ToArray(),
                "application containing the app role");

            AnsiConsole.Clear();

            var selectedApplicationName = AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title("Select the application containing the app role: ")
                    .AddChoices(filteredApplications));

            AnsiConsole.Clear();

            var application = applicationsWithAppRole.Single(x =>
                x.DisplayName.Equals(selectedApplicationName, StringComparison.OrdinalIgnoreCase));

            var oneOf = await AnsiConsole.Status()
                .StartAsync($"Requesting access token from [deepskyblue3_1]Azure[/]...", async ctx =>
                {
                    try
                    {
                        // TODO: Allow user to select which identifier and scope
                        var scope = $"{application.IdentifierUris.First()}/.default";
                        var accessToken = await tokenCredential.GetTokenAsync(new TokenRequestContext([scope]), token);
                        return OneOf<Azure.Core.AccessToken, Exception>.FromT0(accessToken);
                    }
                    catch (Exception e)
                    {
                        return OneOf<Azure.Core.AccessToken, Exception>.FromT1(e);
                    }
                });

            oneOf.Switch(
                accessToken =>
                {
                    try
                    {
                        ClipboardService.SetText(accessToken.Token);
                        AnsiConsole.WriteLine( $"{Environment.NewLine}The access token has been copied to your clipboard.{Environment.NewLine}");
                    } catch
                    {
                        AnsiConsole.WriteLine( $"{Environment.NewLine}Failed to copy the access token to clipboard. Falling back to printing to console:{Environment.NewLine}");
                    }
                    System.Console.WriteLine(accessToken.Token);
                },
                exception => {
                    AnsiConsole.WriteException(exception);
                });

            var generateAnotherToken = AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title("Do you want to generate another access token?")
                    .AddChoices("yes", "no"));

            isRunning = generateAnotherToken == "yes";

            AnsiConsole.Clear();
        }
    }

    private async Task<List<Application>> GetApplicationsAsync(CancellationToken token = default)
    {
        if (_cachedApplications.Any())
        {
            return _cachedApplications.ToList();
        }
        
        var getApplicationsResponse = await AnsiConsole.Progress().StartAsync(async ctx =>
        {
            var progress = ctx.AddTask("Fetching applications");

            var getApplicationsResponse = await graphClient.Applications.GetAsync(cancellationToken: token);

            var applications = new List<Application>();

            var pageIterator = PageIterator<Application, ApplicationCollectionResponse>.CreatePageIterator(
                graphClient,
                getApplicationsResponse,
                (sp) =>
                {
                    progress.Increment(0.5);
                    applications.Add(sp);
                    return Task.FromResult(true);
                });

            await pageIterator.IterateAsync(token);

            while (pageIterator.State != PagingState.Complete)
            {
                await pageIterator.ResumeAsync(token);
            }

            return applications;
        });

        _cachedApplications = getApplicationsResponse.ToArray();
        
        return getApplicationsResponse;
    }

    private string[] FilterWithFuzzySearch(string[] options, string customText)
    {
        var filter = "";
        var table = new Table();
        table.AddColumn("Name");

        ConsoleKeyInfo key;
        do
        {
            table.Rows.Clear();

            FuzzySharp.Process.ExtractTop(filter, options, limit: 10)
                .Select(x => x.Value)
                .ToList()
                .ForEach(x => table.AddRow(new Text(x)));

            AnsiConsole.Clear();
            AnsiConsole.Write(table);
            AnsiConsole.Write(new Rule());
            AnsiConsole.Write(new Text($"Please type in the name of the {customText} to filter the results: " +
                                       filter));

            ProcessKey();

            void ProcessKey()
            {
                key = System.Console.ReadKey(true);

                if (key.Key == ConsoleKey.Backspace && filter.Length > 0)
                {
                    filter = filter.Substring(0, filter.Length - 1);
                }
                else if (key.Key != ConsoleKey.Enter && !char.IsControl(key.KeyChar) || key.Key == ConsoleKey.Spacebar)
                {
                    filter += key.KeyChar;
                }
                else if (key.Key is ConsoleKey.Escape)
                {
                    throw new OperationCanceledException();
                }
                else if (key.Key is not ConsoleKey.Enter)
                {
                    ProcessKey();
                }
            }
        } while (key.Key != ConsoleKey.Enter);

        return FuzzySharp.Process.ExtractTop(filter, options, limit: 10)
            .Select(x => x.Value)
            .ToArray();
    }
}
