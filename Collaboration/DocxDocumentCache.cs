using Microsoft.Extensions.Logging;
using YDotNet.Document;
using YDotNet.Server;
using YDotNet.Server.Storage;

namespace SuperDocPoc.Collaboration;

/// <summary>
/// Custom DocumentCache that uses DocxDocumentContainer instead of the default DocumentContainer
/// to handle DOCX-to-Y.js conversion
/// </summary>
internal sealed class DocxDocumentCache(
    IDocumentStorage documentStorage,
    IDocumentCallback documentCallback,
    IDocumentManager documentManager,
    DocumentManagerOptions options,
    ILogger logger) : IAsyncDisposable
{
    private readonly Dictionary<string, Item> documents = new(StringComparer.Ordinal);
    private readonly SemaphoreSlim slimLock = new(1);

    private sealed class Item
    {
        required public DocxDocumentContainer Document { get; init; }
        public DateTime ValidUntil { get; set; }
    }

    public Func<DateTime> Clock { get; set; } = () => DateTime.UtcNow;

    public async ValueTask DisposeAsync()
    {
        await slimLock.WaitAsync().ConfigureAwait(false);
        try
        {
            foreach (var (_, item) in documents)
            {
                await item.Document.DisposeAsync().ConfigureAwait(false);
            }

            documents.Clear();
        }
        finally
        {
            slimLock.Release();
        }
    }

    public async Task RemoveEvictedItemsAsync()
    {
        // Keep the lock as short as possible.
        var toCleanup = GetItemsToRemove();

        foreach (var document in toCleanup)
        {
            await document.DisposeAsync().ConfigureAwait(false);
        }
    }

    public async Task<T> ApplyUpdateReturnAsync<T>(string documentName, Func<Doc, Task<T>> action)
    {
        var container = await GetAsync(documentName).ConfigureAwait(false);

        return await container.ApplyUpdateReturnAsync(action).ConfigureAwait(false);
    }

    private async Task<DocxDocumentContainer> GetAsync(string documentName)
    {
        await slimLock.WaitAsync().ConfigureAwait(false);
        try
        {
            var now = Clock();

            if (documents.TryGetValue(documentName, out var found))
            {
                found.ValidUntil = now.Add(options.CacheDuration);

                return found.Document;
            }

            var document = new DocxDocumentContainer(
                documentName,
                documentStorage,
                documentCallback,
                documentManager,
                options,
                logger);

            documents[documentName] = new Item
            {
                Document = document,
                ValidUntil = now.Add(options.CacheDuration)
            };

            return document;
        }
        finally
        {
            slimLock.Release();
        }
    }

    private List<DocxDocumentContainer> GetItemsToRemove()
    {
        slimLock.Wait();
        try
        {
            var now = Clock();
            var toRemove = new List<string>();
            var toCleanup = new List<DocxDocumentContainer>();

            foreach (var (key, item) in documents)
            {
                if (item.ValidUntil < now)
                {
                    toRemove.Add(key);
                    toCleanup.Add(item.Document);
                }
            }

            foreach (var key in toRemove)
            {
                documents.Remove(key);
            }

            return toCleanup;
        }
        finally
        {
            slimLock.Release();
        }
    }
}
