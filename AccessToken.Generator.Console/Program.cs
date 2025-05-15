using AccessToken.Generator.Console;
using Spectre.Console;

var cancellationTokenSource = new CancellationTokenSource();

try
{
    var services = new GraphServiceClientBuilder().BuildGraphServices();
    var generator = new AccessTokenGenerator(services.graphServiceClient, services.credential);
    await generator.GenerateAsync(cancellationTokenSource.Token);
}
catch (Exception e)
{
    if (e is OperationCanceledException)
    {
        AnsiConsole.MarkupLine("Operation was canceled by the user.");
        return;
    }

    AnsiConsole.WriteException(e);
    throw;
}
finally
{
    cancellationTokenSource.Dispose();
}