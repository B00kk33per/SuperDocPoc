using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace SuperDocPoc.Converter;

/// <summary>
/// Class to handle unzipping and zipping of docx files
/// </summary>
public class DocxZipper
{
    private readonly bool _debug;
    private readonly List<DocxFile> _files = new();
    private readonly Dictionary<string, byte[]> _fonts = new();

    public DocxZipper(bool debug = false)
    {
        _debug = debug;
    }

    public List<DocxFile> Files => _files;
    public Dictionary<string, byte[]> Fonts => _fonts;

    /// <summary>
    /// Get all docx data from the zipped docx
    /// 
    /// [Content_Types].xml
    /// _rels/.rels
    /// word/document.xml
    /// word/_rels/document.xml.rels
    /// word/footnotes.xml
    /// word/endnotes.xml
    /// word/header1.xml
    /// word/theme/theme1.xml
    /// word/settings.xml
    /// word/styles.xml
    /// word/webSettings.xml
    /// word/fontTable.xml
    /// docProps/core.xml
    /// docProps/app.xml
    /// </summary>
    /// <param name="fileStream">The docx file stream</param>
    /// <returns>List of DocxFile objects containing name and content</returns>
    public async Task<List<DocxFile>> GetDocxDataAsync(Stream fileStream)
    {
        var mediaObjects = new Dictionary<string, string>(); // Local variable, not stored in class
        var validTypes = new[] { "xml", "rels" };

        using var archive = new ZipArchive(fileStream, ZipArchiveMode.Read, leaveOpen: true);

        foreach (var entry in archive.Entries)
        {
            if (validTypes.Any(validType => entry.FullName.EndsWith(validType)))
            {
                using var entryStream = entry.Open();
                using var reader = new StreamReader(entryStream);
                var content = await reader.ReadToEndAsync();

                _files.Add(new DocxFile
                {
                    Name = entry.FullName,
                    Content = content
                });
            }
            else if (entry.FullName.StartsWith("word/media") && entry.FullName != "word/media/")
            {
                // Server environment - convert buffer to base64 (equivalent to isNode = true)
                using var entryStream = entry.Open();
                using var memoryStream = new MemoryStream();
                await entryStream.CopyToAsync(memoryStream);
                var buffer = memoryStream.ToArray();
                var fileBase64 = Convert.ToBase64String(buffer);
                mediaObjects[entry.FullName] = fileBase64; // Local variable, just like JavaScript
            }
            else if (entry.FullName.StartsWith("word/fonts") && entry.FullName != "word/fonts/")
            {
                using var entryStream = entry.Open();
                using var memoryStream = new MemoryStream();
                await entryStream.CopyToAsync(memoryStream);
                _fonts[entry.FullName] = memoryStream.ToArray();
            }
        }

        return _files;
    }

    /// <summary>
    /// Get file extension from filename
    /// </summary>
    private static string GetFileExtension(string fileName)
    {
        return Path.GetExtension(fileName).TrimStart('.');
    }

    /// <summary>
    /// Update [Content_Types].xml with extensions of new Image annotations
    /// </summary>
    public async Task<string> UpdateContentTypesAsync(
        List<DocxFile> docxFiles,
        Dictionary<string, byte[]> media,
        Dictionary<string, string> additionalFiles = null)
    {
        var newMediaTypes = media.Keys.Select(GetFileExtension).ToList();

        const string contentTypesPath = "[Content_Types].xml";
        var contentTypesFile = docxFiles.FirstOrDefault(f => f.Name == contentTypesPath);
        if (contentTypesFile == null)
            throw new InvalidOperationException("Content_Types.xml not found");

        var contentTypesXml = contentTypesFile.Content;
        var typesStringBuilder = new StringBuilder();

        var defaultMediaTypes = GetContentTypesFromXml(contentTypesXml);

        // Update media types in content types
        var seenTypes = new HashSet<string>();
        foreach (var type in newMediaTypes)
        {
            // Current extension already presented in Content_Types
            if (defaultMediaTypes.Contains(type) || seenTypes.Contains(type))
                continue;

            var newContentType = $"<Default Extension=\"{type}\" ContentType=\"image/{type}\"/>";
            typesStringBuilder.Append(newContentType);
            seenTypes.Add(type);
        }

        // Parse XML for comments and overrides
        var xmlDoc = XDocument.Parse(contentTypesXml);
        var typesElement = xmlDoc.Root;
        var namespaceUri = typesElement?.Name.Namespace ?? XNamespace.None;

        // Check for existing overrides
        var overrides = typesElement?.Elements(namespaceUri + "Override")?.ToList() ?? new List<XElement>();

        var hasComments = overrides.Any(el => el.Attribute("PartName")?.Value == "/word/comments.xml");
        var hasCommentsExtended = overrides.Any(el => el.Attribute("PartName")?.Value == "/word/commentsExtended.xml");
        var hasCommentsIds = overrides.Any(el => el.Attribute("PartName")?.Value == "/word/commentsIds.xml");
        var hasCommentsExtensible = overrides.Any(el => el.Attribute("PartName")?.Value == "/word/commentsExtensible.xml");

        // Add comments-related overrides if files exist
        if (additionalFiles?.ContainsKey("word/comments.xml") == true && !hasComments)
        {
            typesStringBuilder.Append("<Override PartName=\"/word/comments.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml\" />");
        }

        if (additionalFiles?.ContainsKey("word/commentsExtended.xml") == true && !hasCommentsExtended)
        {
            typesStringBuilder.Append("<Override PartName=\"/word/commentsExtended.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml\" />");
        }

        if (additionalFiles?.ContainsKey("word/commentsIds.xml") == true && !hasCommentsIds)
        {
            typesStringBuilder.Append("<Override PartName=\"/word/commentsIds.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml\" />");
        }

        if (additionalFiles?.ContainsKey("word/commentsExtensible.xml") == true && !hasCommentsExtensible)
        {
            typesStringBuilder.Append("<Override PartName=\"/word/commentsExtensible.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml\" />");
        }

        // Add header/footer overrides
        if (additionalFiles != null)
        {
            foreach (var fileName in additionalFiles.Keys)
            {
                if (!fileName.Contains("header") && !fileName.Contains("footer")) continue;

                var hasExtensible = overrides.Any(el => el.Attribute("PartName")?.Value == $"/{fileName}");
                var type = fileName.Contains("header") ? "header" : "footer";

                if (!hasExtensible)
                {
                    typesStringBuilder.Append($"<Override PartName=\"/{fileName}\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.{type}+xml\"/>");
                }
            }
        }

        const string beginningString = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">";
        var updatedContentTypesXml = contentTypesXml.Replace(beginningString, $"{beginningString}{typesStringBuilder}");

        return updatedContentTypesXml;
    }

    /// <summary>
    /// Create a new ZIP archive with updated documents
    /// </summary>
    public async Task<byte[]> UpdateZipAsync(
        List<DocxFile> docxFiles,
        Dictionary<string, string> updatedDocs,
        Stream originalDocxFile = null,
        Dictionary<string, byte[]> media = null,
        Dictionary<string, byte[]> fonts = null)
    {
        media ??= new Dictionary<string, byte[]>();
        fonts ??= new Dictionary<string, byte[]>();

        using var memoryStream = new MemoryStream();

        if (originalDocxFile != null)
        {
            await ExportFromOriginalFileAsync(originalDocxFile, updatedDocs, media, memoryStream);
        }
        else
        {
            await ExportFromCollaborativeDocxAsync(docxFiles, updatedDocs, media, fonts, memoryStream);
        }

        return memoryStream.ToArray();
    }

    /// <summary>
    /// Export from collaborative docx files
    /// </summary>
    private async Task ExportFromCollaborativeDocxAsync(
        List<DocxFile> docxFiles,
        Dictionary<string, string> updatedDocs,
        Dictionary<string, byte[]> media,
        Dictionary<string, byte[]> fonts,
        Stream outputStream)
    {
        using var archive = new ZipArchive(outputStream, ZipArchiveMode.Create, leaveOpen: true);

        // Add original files
        foreach (var file in docxFiles)
        {
            var entry = archive.CreateEntry(file.Name);
            using var entryStream = entry.Open();
            using var writer = new StreamWriter(entryStream);
            await writer.WriteAsync(file.Content);
        }

        // Replace updated docs
        foreach (var kvp in updatedDocs)
        {
            var existingEntry = archive.Entries.FirstOrDefault(e => e.FullName == kvp.Key);
            existingEntry?.Delete();

            var entry = archive.CreateEntry(kvp.Key);
            using var entryStream = entry.Open();
            using var writer = new StreamWriter(entryStream);
            await writer.WriteAsync(kvp.Value);
        }

        // Add media files
        foreach (var kvp in media)
        {
            var entry = archive.CreateEntry($"word/media/{kvp.Key}");
            using var entryStream = entry.Open();
            await entryStream.WriteAsync(kvp.Value);
        }

        // Add font files
        foreach (var kvp in fonts)
        {
            var entry = archive.CreateEntry(kvp.Key);
            using var entryStream = entry.Open();
            await entryStream.WriteAsync(kvp.Value);
        }

        // Update content types
        var updatedContentTypes = await UpdateContentTypesAsync(docxFiles, media, updatedDocs);
        var contentTypesEntry = archive.Entries.FirstOrDefault(e => e.FullName == "[Content_Types].xml");
        contentTypesEntry?.Delete();

        var newContentTypesEntry = archive.CreateEntry("[Content_Types].xml");
        using var contentTypesStream = newContentTypesEntry.Open();
        using var contentTypesWriter = new StreamWriter(contentTypesStream);
        await contentTypesWriter.WriteAsync(updatedContentTypes);
    }

    /// <summary>
    /// Export from original file
    /// </summary>
    private async Task ExportFromOriginalFileAsync(
        Stream originalDocxFile,
        Dictionary<string, string> updatedDocs,
        Dictionary<string, byte[]> media,
        Stream outputStream)
    {
        originalDocxFile.Position = 0;

        using var inputArchive = new ZipArchive(originalDocxFile, ZipArchiveMode.Read, leaveOpen: true);
        using var outputArchive = new ZipArchive(outputStream, ZipArchiveMode.Create, leaveOpen: true);

        // Copy all entries from original
        foreach (var entry in inputArchive.Entries)
        {
            var newEntry = outputArchive.CreateEntry(entry.FullName);
            using var originalStream = entry.Open();
            using var newStream = newEntry.Open();
            await originalStream.CopyToAsync(newStream);
        }

        // Replace updated docs
        foreach (var kvp in updatedDocs)
        {
            var existingEntry = outputArchive.Entries.FirstOrDefault(e => e.FullName == kvp.Key);
            existingEntry?.Delete();

            var entry = outputArchive.CreateEntry(kvp.Key);
            using var entryStream = entry.Open();
            using var writer = new StreamWriter(entryStream);
            await writer.WriteAsync(kvp.Value);
        }

        // Add media files
        foreach (var kvp in media)
        {
            var entry = outputArchive.CreateEntry($"word/media/{kvp.Key}");
            using var entryStream = entry.Open();
            await entryStream.WriteAsync(kvp.Value);
        }

        // Update content types
        var docxFiles = new List<DocxFile>();
        // Note: You'll need to extract docxFiles from the archive for UpdateContentTypesAsync
        // This is a simplified version - you might need to adjust based on your needs
    }

    /// <summary>
    /// Get content types from XML - placeholder for the helper function
    /// </summary>
    private static string[] GetContentTypesFromXml(string contentTypesXml)
    {
        var xmlDoc = XDocument.Parse(contentTypesXml);
        var namespaceUri = xmlDoc.Root?.Name.Namespace ?? XNamespace.None;

        return xmlDoc.Root?
            .Elements(namespaceUri + "Default")?
            .Select(el => el.Attribute("Extension")?.Value ?? string.Empty)
            .Where(ext => !string.IsNullOrEmpty(ext))
            .ToArray() ?? [];
    }
}

/// <summary>
/// Represents a file within a DOCX archive
/// </summary>
public class DocxFile
{
    public string Name { get; set; }
    public string Content { get; set; }
}
