
using System.Collections.Concurrent;
using Microsoft.Extensions.Logging;
using YDotNet.Server.Storage;

namespace SuperDocPoc.Collaboration;

public class FileSystemDocumentStorage(string rootPath, ILogger<FileSystemDocumentStorage> log) : IDocumentStorage
{
    private readonly ConcurrentDictionary<string, byte[]> docs = new(StringComparer.Ordinal);
    
    public async ValueTask<byte[]?> GetDocAsync(string name, CancellationToken ct = default)
    {
        log.LogInformation("Retrieving document: {DocumentName}", name);

        if (docs.TryGetValue(name, out var doc))
            return await new ValueTask<byte[]>(doc);

        var path = GetDocumentPath(name);
        if (!File.Exists(path))
        {
            log.LogWarning("Document file not found: {FilePath}", path);
            path = GetDocumentPath("sample.docx");
            //return null;
        }

        try
        {
            var fileData = await File.ReadAllBytesAsync(path, ct);
            docs.TryAdd(name, fileData); // Cache the loaded document
            log.LogInformation("Successfully loaded document {DocumentName} from {FilePath}", name, path);
            return fileData;
        }
        catch (Exception ex)
        {
            log.LogError(ex, "Failed to read document file: {FilePath}", path);
            throw;
        }
    }

    public async ValueTask StoreDocAsync(string name, byte[] doc, CancellationToken ct = default)
    {
        log.LogInformation("Storing document: {DocumentName}", name);

        var path = GetDocumentPath(name);
        var directory = Path.GetDirectoryName(path);
        
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        try
        {
            await File.WriteAllBytesAsync(path, doc, ct);
            docs.AddOrUpdate(name, doc, (_, _) => doc); // Update cache
            log.LogInformation("Successfully stored document {DocumentName} to {FilePath}", name, path);
        }
        catch (Exception ex)
        {
            log.LogError(ex, "Failed to store document file: {FilePath}", path);
            throw;
        }
    }

    private string GetDocumentPath(string name)
    {
        // Sanitize the document name to prevent directory traversal
        var sanitizedName = Path.GetFileName(name);
        if (string.IsNullOrEmpty(sanitizedName))
        {
            throw new ArgumentException("Invalid document name", nameof(name));
        }

        // Add .docx extension if not present and if it's likely a DOCX file
        if (!sanitizedName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) && 
            !sanitizedName.EndsWith(".yjs", StringComparison.OrdinalIgnoreCase))
        {
            sanitizedName += ".docx";
        }

        return Path.Combine(rootPath, sanitizedName);
    }
}
