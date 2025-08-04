using StepWise.Prose.Model;

using SuperDocPoc.Converter;

using YDotNet.Document;

namespace SuperDocPoc.Collaboration;
/// <summary>
/// Integration service that combines DocxZipper and SuperConverterNet
/// for complete DOCX to ProseMirror conversion pipeline
/// </summary>
public class DocxToProseMirrorService
{
    private readonly DocxZipper _docxZipper;
    private readonly SuperConverterNet _superConverter;

    public DocxToProseMirrorService()
    {
        _docxZipper = new DocxZipper();
    }

    /// <summary>
    /// Complete pipeline: DOCX file → XML extraction → ProseMirror document
    /// </summary>
    /// <param name="docxFileStream">DOCX file stream</param>
    /// <returns>Conversion result with ProseMirror document and metadata</returns>
    public async Task<DocxConversionResult> ConvertDocxToProseMirrorAsync(Stream docxFileStream)
    {
        try
        {
            // Step 1: Extract DOCX content using DocxZipper
            var extractionResult = await ExtractDocxContentAsync(docxFileStream);

            // Step 2: Create SuperConverterNet with extracted content
            var superConverter = new SuperConverterNet(
                docxXmlFiles: extractionResult.XmlFiles,
                mediaFiles: extractionResult.MediaFiles,
                fonts: extractionResult.Fonts,
                debug: true // Enable for development
            );

            // Step 3: Convert to ProseMirror document
            var proseMirrorDoc = await superConverter.CreateProseMirrorDocumentAsync();

            // Step 4: Return complete result
            return new DocxConversionResult
            {
                Success = true,
                ProseMirrorDocument = proseMirrorDoc,
                ProseMirrorJson = proseMirrorDoc.ToJSON(),
                XmlFiles = extractionResult.XmlFiles,
                MediaFiles = extractionResult.Media,
                MediaFilesBase64 = extractionResult.MediaFiles,
                Fonts = extractionResult.Fonts,
                SuperConverter = superConverter
            };
        }
        catch (Exception ex)
        {
            return new DocxConversionResult
            {
                Success = false,
                Error = ex.Message,
                Exception = ex
            };
        }
    }

    /// <summary>
    /// Extract DOCX content using DocxZipper
    /// </summary>
    /// <param name="docxFileStream">DOCX file stream</param>
    /// <returns>Extraction result</returns>
    private async Task<DocxExtractionResult> ExtractDocxContentAsync(Stream docxFileStream)
    {
        // Extract all content from DOCX
        var docxFiles = await _docxZipper.GetDocxDataAsync(docxFileStream);

        // Convert DocxFile list to XML files dictionary
        var xmlFiles = docxFiles.ToDictionary(f => f.Name, f => f.Content);

        // Get media files as base64 (server environment)
        var mediaFilesBase64 = _docxZipper.MediaFiles.ToDictionary(
            kvp => kvp.Key,
            kvp => (object)kvp.Value
        );

        // Get media information (same as MediaFiles in simplified version)
        var media = _docxZipper.MediaFiles.ToDictionary(
            kvp => kvp.Key,
            kvp => kvp.Value
        );

        // Get fonts as byte arrays
        var fonts = _docxZipper.Fonts.ToDictionary(
            kvp => kvp.Key,
            kvp => (object)kvp.Value
        );

        return new DocxExtractionResult
        {
            XmlFiles = xmlFiles,
            MediaFiles = mediaFilesBase64,
            Media = media,
            Fonts = fonts
        };
    }

    /// <summary>
    /// Prepare data for Y.js collaboration using ydotnet
    /// </summary>
    /// <param name="conversionResult">DOCX conversion result</param>
    /// <returns>Collaboration data</returns>
    public CollaborationData PrepareCollaborationData(DocxConversionResult conversionResult)
    {
        if (!conversionResult.Success)
        {
            throw new InvalidOperationException("Cannot prepare collaboration data from failed conversion");
        }

        return new CollaborationData
        {
            // ProseMirror document JSON for Y.js fragment
            DocumentJson = conversionResult.ProseMirrorJson,

            // XML files from DOCX (for docx metaMap key)
            XmlFiles = conversionResult.XmlFiles,

            // Media files as base64 for Y.js media map
            MediaMap = conversionResult.MediaFilesBase64,

            // Metadata for Y.js meta map
            MetaData = new Dictionary<string, object>
            {
                ["version"] = "1.0.0", // SuperDoc version
                ["created"] = DateTime.UtcNow.ToString("O"),
                ["documentId"] = Guid.NewGuid().ToString()
            }
        };
    }

    /// <summary>
    /// Convert DOCX directly to Y.js document for collaboration
    /// This method combines DOCX → ProseMirror → Y.js conversion in one call
    /// </summary>
    /// <param name="docxFileStream">DOCX file stream</param>
    /// <returns>Y.js document ready for collaboration</returns>
    public async Task<YjsConversionResult> ConvertDocxToYjsAsync(Stream docxFileStream)
    {
        try
        {
            // Use the DocxToYdocService for the complete conversion
            var ydoc = await DocxToYdocService.ConvertDocxToYDocAsync(docxFileStream);

            // Get binary update using transaction
            byte[] binaryUpdate;
            using (var transaction = ydoc.ReadTransaction())
            {
                binaryUpdate = transaction.StateDiffV1(null); // null state vector means get full state
            }

            return new YjsConversionResult
            {
                Success = true,
                YjsDocument = ydoc,
                BinaryUpdate = binaryUpdate
            };
        }
        catch (Exception ex)
        {
            return new YjsConversionResult
            {
                Success = false,
                Error = ex.Message,
                Exception = ex
            };
        }
    }

    /// <summary>
    /// Example of how to integrate with ydotnet
    /// This shows the pattern for storing in Y.js document
    /// </summary>
    /// <param name="collaborationData">Prepared collaboration data</param>
    /// <returns>Y.js integration example</returns>
    public string GetYdotnetIntegrationExample(CollaborationData collaborationData)
    {
        return $@"
// Example C# code for ydotnet integration:

using YDotNet.Document;
using YDotNet.Document.Types.Maps;

// Create Y.js document
var ydoc = new Doc();

// Store ProseMirror document
var prosemirrorMap = ydoc.Map(""prosemirror"");
using (var transaction = ydoc.WriteTransaction())
{{
    prosemirrorMap.Insert(transaction, ""document"", Input.String(JsonSerializer.Serialize(collaborationData.DocumentJson)));
}}

// Store media files
var mediaMap = ydoc.Map(""media"");
using (var transaction = ydoc.WriteTransaction())
{{
    foreach (var media in collaborationData.MediaMap)
    {{
        mediaMap.Insert(transaction, media.Key, Input.String(media.Value.ToString()));
    }}
}}

// Store metadata
var metaMap = ydoc.Map(""meta"");
using (var transaction = ydoc.WriteTransaction())
{{
    foreach (var meta in collaborationData.MetaData)
    {{
        metaMap.Insert(transaction, meta.Key, Input.String(meta.Value.ToString()));
    }}
}}

// Get binary update for collaboration
byte[] binaryUpdate;
using (var transaction = ydoc.ReadTransaction())
{{
    binaryUpdate = transaction.StateDiffV1(null); // Get full state as binary
}}
";
    }
}

#region Result Classes

/// <summary>
/// Result of DOCX to ProseMirror conversion
/// </summary>
public class DocxConversionResult
{
    public bool Success { get; set; }
    public Node ProseMirrorDocument { get; set; }
    public object ProseMirrorJson { get; set; }
    public Dictionary<string, string> XmlFiles { get; set; }
    public Dictionary<string, string> MediaFiles { get; set; }
    public Dictionary<string, object> MediaFilesBase64 { get; set; }
    public Dictionary<string, object> Fonts { get; set; }
    public SuperConverterNet SuperConverter { get; set; }
    public string Error { get; set; }
    public Exception Exception { get; set; }
}

/// <summary>
/// Result of DOCX extraction
/// </summary>
public class DocxExtractionResult
{
    public Dictionary<string, string> XmlFiles { get; set; }
    public Dictionary<string, string> Media { get; set; }
    public Dictionary<string, object> MediaFiles { get; set; }
    public Dictionary<string, object> Fonts { get; set; }
}

/// <summary>
/// Data prepared for Y.js collaboration
/// </summary>
public class CollaborationData
{
    public object DocumentJson { get; set; }
    public Dictionary<string, string> XmlFiles { get; set; }
    public Dictionary<string, object> MediaMap { get; set; }
    public Dictionary<string, object> MetaData { get; set; }
}

/// <summary>
/// Result of DOCX to Y.js conversion
/// </summary>
public class YjsConversionResult
{
    public bool Success { get; set; }
    public Doc YjsDocument { get; set; }
    public byte[] BinaryUpdate { get; set; }
    public string Error { get; set; }
    public Exception Exception { get; set; }
}

#endregion
