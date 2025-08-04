using Microsoft.Extensions.Logging;
using YDotNet.Document;
using YDotNet.Server;
using YDotNet.Server.Storage;
using SuperDocPoc.Converter;

namespace SuperDocPoc.Collaboration;

/// <summary>
/// Custom DocumentContainer that loads DOCX files and converts them to Y.js documents
/// using DocxToYdocService instead of applying raw Y.js binary data
/// </summary>
internal sealed class DocxDocumentContainer
{
    private readonly DocumentManagerOptions options;
    private readonly ILogger logger;
    private readonly DocxDelayedWriter delayedWriter;
    private readonly string documentName;
    private readonly IDocumentStorage documentStorage;
    private readonly Task<Doc> loadingTask;
    private readonly SemaphoreSlim slimLock = new(1);

    public string Name => documentName;

    public DocxDocumentContainer(
        string documentName,
        IDocumentStorage documentStorage,
        IDocumentCallback documentCallback,
        IDocumentManager documentManager,
        DocumentManagerOptions options,
        ILogger logger)
    {
        this.documentName = documentName;
        this.documentStorage = documentStorage;
        this.options = options;
        this.logger = logger;

        delayedWriter = new DocxDelayedWriter(options.StoreDebounce, options.MaxWriteTimeInterval, WriteAsync);

        loadingTask = LoadInternalAsync(documentCallback, documentManager, logger);
    }

    private async Task<Doc> LoadInternalAsync(IDocumentCallback documentCallback, IDocumentManager documentManager, ILogger logger)
    {
        var doc = await LoadCoreAsync().ConfigureAwait(false);

        await documentCallback.OnDocumentLoadedAsync(new DocumentLoadEvent
        {
            Document = doc,
            Context = new DocumentContext(documentName, 0),
            Source = documentManager,
        }).ConfigureAwait(false);

        doc.ObserveUpdatesV1(e =>
        {
            logger.LogDebug("Document {name} updated.", documentName);
            delayedWriter.Ping();
        });

        return doc;
    }

    private async Task<Doc> LoadCoreAsync()
    {
        var documentData = await documentStorage.GetDocAsync(documentName).ConfigureAwait(false);

        if (documentData != null)
        {
            // Check if this is DOCX data (starts with PK signature) or Y.js binary data
            if (IsDocxFile(documentData))
            {
                logger.LogInformation("Loading DOCX file {documentName} and converting to Y.js", documentName);
                
                // Convert DOCX to Y.js document using our service
                using var stream = new MemoryStream(documentData);
                var ydoc = await DocxToYdocService.ConvertDocxToYDocAsync(stream);
                
                logger.LogInformation("Successfully converted DOCX {documentName} to Y.js document", documentName);
                return ydoc;
            }
            else
            {
                logger.LogInformation("Loading existing Y.js document {documentName}", documentName);
                
                // This is already Y.js binary data, apply it normally
                var document = new Doc();
                using (var transaction = document.WriteTransaction())
                {
                    if (transaction == null)
                    {
                        throw new InvalidOperationException("Transaction cannot be acquired.");
                    }
                    transaction.ApplyV1(documentData);
                }
                return document;
            }
        }

        if (options.AutoCreateDocument)
        {
            logger.LogInformation("Creating new empty document {documentName}", documentName);
            return new Doc();
        }

        throw new InvalidOperationException($"Document {documentName} does not exist yet.");
    }

    /// <summary>
    /// Check if the byte array represents a DOCX file (ZIP format with PK signature)
    /// </summary>
    private static bool IsDocxFile(byte[] data)
    {
        return data.Length >= 4 && 
               data[0] == 0x50 && // 'P'
               data[1] == 0x4B && // 'K'
               (data[2] == 0x03 || data[2] == 0x05 || data[2] == 0x07) &&
               (data[3] == 0x04 || data[3] == 0x06 || data[3] == 0x08);
    }

    public async Task DisposeAsync()
    {
        await delayedWriter.FlushAsync().ConfigureAwait(false);
    }

    public async Task<T> ApplyUpdateReturnAsync<T>(Func<Doc, Task<T>> action)
    {
        var document = await loadingTask.ConfigureAwait(false);

        // This is the only option to get access to the document to prevent concurrency issues.
        await slimLock.WaitAsync().ConfigureAwait(false);
        try
        {
            return await action(document).ConfigureAwait(false);
        }
        finally
        {
            slimLock.Release();
        }
    }

    private async Task WriteAsync()
    {
        // var document = await loadingTask.ConfigureAwait(false);

        // logger.LogDebug("Document {documentName} will be saved.", documentName);
        // try
        // {
        //     // All the writes are thread safe itself, but they have to be synchronized with a write.
        //     var state = GetStateLocked(document);

        //     await documentStorage.StoreDocAsync(documentName, state).ConfigureAwait(false);

        //     logger.LogDebug("Document {documentName} with size {size} has been saved.", documentName, state.Length);
        // }
        // catch (Exception ex)
        // {
        //     logger.LogError(ex, "Document {documentName} could not be saved.", documentName);
        // }
    }

    private byte[] GetStateLocked(Doc document)
    {
        slimLock.Wait();
        try
        {
            using var transaction = document.ReadTransaction();
            return transaction.StateDiffV1(stateVector: null)!;
        }
        finally
        {
            slimLock.Release();
        }
    }
}
