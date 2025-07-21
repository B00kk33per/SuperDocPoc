using SuperDocPoc.Collaboration;

using YDotNet.Server;
using YDotNet.Server.Storage;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllersWithViews();

var collaborationHooks = new MyCollaborationHooks();

// Register YDotNet services with WebSocket support
builder.Services.AddYDotNet()
    .AutoCleanup()
    .AddCallback<Callback>() // Replace with your callback if needed
    .AddWebSockets(options => {
        options.OnAuthenticateAsync = async (httpContext, docContext) =>
        {
            // Call your custom authentication logic
            var isAuthenticated = await collaborationHooks.AuthenticateAsync(httpContext);
            if (!isAuthenticated)
            {
                httpContext.Response.StatusCode = StatusCodes.Status401Unauthorized;
                await httpContext.Response.CompleteAsync();
            }
        };
    });

builder.Services.AddSingleton<IDocumentStorage>(provider =>
    new FileSystemDocumentStorage(Path.Combine(builder.Environment.WebRootPath, "docs"), provider.GetRequiredService<ILogger<FileSystemDocumentStorage>>()));

// Optional: Configure document manager options (cache, etc.)
builder.Services.Configure<DocumentManagerOptions>(options =>
{
    options.CacheDuration = TimeSpan.FromSeconds(10);
    options.StoreDebounce = TimeSpan.FromMilliseconds(500);
});

var app = builder.Build();

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseWebSockets();
app.UseRouting();
app.UseAuthorization();

// Map default MVC route
app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.Map("/collaboration", branch =>
{
    branch.UseYDotnetWebSockets();
});

app.Run();