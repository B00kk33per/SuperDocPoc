
using System.Collections.Concurrent;

using YDotNet.Server.Storage;

namespace SuperDocPoc.Collaboration;

public class FileSystemDocumentStorage(string rootPath, ILogger<FileSystemDocumentStorage> log) : IDocumentStorage
{
    private readonly ConcurrentDictionary<string, byte[]> docs = new(StringComparer.Ordinal);
    public async ValueTask<byte[]?> GetDocAsync(string name, CancellationToken ct = default)
    {
        log.LogInformation("Retrieving document: {DocumentName}", name);

        if(docs.TryGetValue(name, out var doc))
            return await new ValueTask<byte[]>(doc);

        var path = GetDocumentPath(name);
        if (!File.Exists(path))
        {
            log.LogWarning("Document not found: {DocumentPath}", path);
            path = GetDocumentPath("sample");
        }

        return await File.ReadAllBytesAsync(path, ct);

    }

    public ValueTask StoreDocAsync(string name, byte[] doc, CancellationToken ct = default)
    {
        log.LogInformation("Storing document: {DocumentName}", name);
        docs[name] = doc;

        return default;
    }

    private string GetDocumentPath(string documentId)
    {
        // Normalizes paths like "demo-doc" to "demo-doc.docx"
        var fileName = $"{documentId}.docx";
        return Path.Combine(rootPath, fileName);
    }

}
