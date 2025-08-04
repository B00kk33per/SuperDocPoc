using StepWise.Prose.Collections;
using StepWise.Prose.Model;
using System.Xml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace SuperDocPoc.Converter;
/// <summary>
/// .NET port of SuperConverter.js for converting DOCX XML to ProseMirror JSON format
/// Uses prosemirror-dotnet for native ProseMirror structure creation
/// </summary>
public class SuperConverterNet
{
    #region Static Mappings (from JavaScript)

    private static readonly Dictionary<string, string> AllowedElements = new()
    {
        ["w:document"] = "doc",
        ["w:body"] = "body",
        ["w:p"] = "paragraph",
        ["w:r"] = "run",
        ["w:t"] = "text",
        ["w:delText"] = "text",
        ["w:br"] = "lineBreak",
        ["w:tbl"] = "table",
        ["w:tr"] = "tableRow",
        ["w:tc"] = "tableCell",
        ["w:drawing"] = "drawing",
        ["w:bookmarkStart"] = "bookmarkStart",

        // Formatting only
        ["w:sectPr"] = "sectionProperties",
        ["w:rPr"] = "runProperties",

        // Comments
        ["w:commentRangeStart"] = "commentRangeStart",
        ["w:commentRangeEnd"] = "commentRangeEnd",
        ["w:commentReference"] = "commentReference"
    };

    private static readonly List<MarkTypeMapping> MarkTypes = new()
        {
            new() { Name = "w:b", Type = "bold" },
            new() { Name = "w:bCs", Type = "bold" },
            new() { Name = "w:i", Type = "italic" },
            new() { Name = "w:iCs", Type = "italic" },
            new() { Name = "w:u", Type = "underline", Mark = "underline", Property = "underlineType" },
            new() { Name = "w:strike", Type = "strike" },
            new() { Name = "w:color", Type = "color", Mark = "textStyle", Property = "color" },
            new() { Name = "w:sz", Type = "fontSize", Mark = "textStyle", Property = "fontSize" },
            new() { Name = "w:szCs", Type = "fontSize", Mark = "textStyle", Property = "fontSize" },
            new() { Name = "w:rFonts", Type = "fontFamily", Mark = "textStyle", Property = "fontFamily" },
            new() { Name = "w:jc", Type = "textAlign", Mark = "textStyle", Property = "textAlign" },
            new() { Name = "w:ind", Type = "textIndent", Mark = "textStyle", Property = "textIndent" },
            new() { Name = "w:spacing", Type = "lineHeight", Mark = "textStyle", Property = "lineHeight" },
            new() { Name = "link", Type = "link", Mark = "link", Property = "href" },
            new() { Name = "w:highlight", Type = "highlight", Mark = "highlight", Property = "color" },
            new() { Name = "w:shd", Type = "highlight", Mark = "highlight", Property = "color" }
        };

    private static readonly Dictionary<string, string> PropertyTypes = new()
    {
        ["w:pPr"] = "paragraphProperties",
        ["w:rPr"] = "runProperties",
        ["w:sectPr"] = "sectionProperties",
        ["w:numPr"] = "numberingProperties",
        ["w:tcPr"] = "tableCellProperties"
    };

    #endregion

    #region Properties

    public bool Debug { get; set; }
    public Dictionary<string, JObject> ConvertedXml { get; private set; } = new();
    public Dictionary<string, object> Media { get; private set; } = new();
    public Dictionary<string, object> Fonts { get; private set; } = new();
    public Dictionary<string, object> AddedMedia { get; private set; } = new();
    public List<object> Comments { get; private set; } = new();
    public JObject InitialJson { get; private set; }
    public XmlDeclaration Declaration { get; private set; }
    public object PageStyles { get; private set; }
    public object Numbering { get; private set; }
    public List<object> SavedTagsToRestore { get; private set; } = new();
    public string DocumentInternalId { get; private set; }

    // Headers and footers
    public Dictionary<string, object> Headers { get; private set; } = new();
    public Dictionary<string, string> HeaderIds { get; private set; } = new()
    {
        ["default"] = null,
        ["even"] = null,
        ["odd"] = null,
        ["first"] = null
    };

    public Dictionary<string, object> Footers { get; private set; } = new();
    public Dictionary<string, string> FooterIds { get; private set; } = new()
    {
        ["default"] = null,
        ["even"] = null,
        ["odd"] = null,
        ["first"] = null
    };

    private readonly Schema _proseMirrorSchema;

    #endregion

    #region Constructor

    /// <summary>
    /// Initialize SuperConverterNet with DOCX XML files and ProseMirror schema
    /// </summary>
    /// <param name="docxXmlFiles">Dictionary of XML files from DocxZipper</param>
    /// <param name="mediaFiles">Media files as base64</param>
    /// <param name="fonts">Font files</param>
    /// <param name="proseMirrorSchema">ProseMirror schema for document creation</param>
    /// <param name="debug">Enable debug logging</param>
    public SuperConverterNet(
        Dictionary<string, string> docxXmlFiles,
        Dictionary<string, object> mediaFiles = null,
        Dictionary<string, object> fonts = null,
        Schema proseMirrorSchema = null,
        bool debug = false)
    {
        Debug = debug;
        Media = mediaFiles ?? new Dictionary<string, object>();
        Fonts = fonts ?? new Dictionary<string, object>();
        _proseMirrorSchema = proseMirrorSchema ?? CreateDefaultSchema();

        if (docxXmlFiles?.Any() == true)
        {
            ParseFromXml(docxXmlFiles);
        }
    }

    #endregion

    #region Core Methods

    /// <summary>
    /// Parse XML files and convert to JSON structure
    /// </summary>
    /// <param name="docxXmlFiles">Dictionary of XML file content</param>
    private void ParseFromXml(Dictionary<string, string> docxXmlFiles)
    {
        foreach (var file in docxXmlFiles)
        {
            try
            {
                ConvertedXml[file.Key] = ParseXmlToJson(file.Value);

                // Store document attributes for main document
                if (file.Key == "word/document.xml")
                {
                    // With Newtonsoft.Json, the structure is different
                    // No need to check for "elements" array
                    if (Debug)
                    {
                        Console.WriteLine($"Parsed {file.Key} successfully");
                    }
                }
            }
            catch (Exception ex)
            {
                if (Debug)
                {
                    Console.WriteLine($"Error parsing {file.Key}: {ex.Message}");
                }

                // Skip non-critical files that have XML parsing issues
                // Only fail for critical document files
                if (IsCriticalDocumentFile(file.Key))
                {
                    throw new InvalidOperationException($"Failed to parse critical XML file {file.Key}", ex);
                }
                else
                {
                    // Log warning but continue processing
                    Console.WriteLine($"Warning: Skipping problematic XML file {file.Key} - {ex.Message}");
                    continue;
                }
            }
        }

        InitialJson = ConvertedXml.ContainsKey("word/document.xml") ? ConvertedXml["word/document.xml"] : null;

        if (Debug && InitialJson != null)
        {
            Console.WriteLine("Successfully parsed word/document.xml");
            Console.WriteLine($"InitialJson root properties: {string.Join(", ", InitialJson.Properties().Select(p => p.Name))}");
        }
    }

    /// <summary>
    /// Determine if a file is critical for document processing
    /// </summary>
    /// <param name="fileName">File name/path</param>
    /// <returns>True if file is critical</returns>
    private bool IsCriticalDocumentFile(string fileName)
    {
        var criticalFiles = new[]
        {
                "word/document.xml",
                "[Content_Types].xml",
                "_rels/.rels",
                "word/_rels/document.xml.rels"
            };

        return criticalFiles.Contains(fileName);
    }

    /// <summary>
    /// Convert XML string to JSON structure (equivalent to xml-js)
    /// </summary>
    /// <param name="xml">XML content</param>
    /// <returns>JSON representation of XML</returns>
    private JObject ParseXmlToJson(string xml)
    {
        try
        {
            // Parse XML
            var xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(xml);

            // Convert to JSON using Newtonsoft.Json (similar to xml2js)
            var jsonString = JsonConvert.SerializeXmlNode(xmlDoc, Newtonsoft.Json.Formatting.None, omitRootObject: false);

            // Parse the JSON string to JObject
            return JObject.Parse(jsonString);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to parse XML: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Main method to create ProseMirror document from DOCX XML
    /// </summary>
    /// <returns>ProseMirror Node representing the document</returns>
    public async Task<Node> CreateProseMirrorDocumentAsync()
    {
        if (InitialJson == null)
        {
            throw new InvalidOperationException("No document XML found. Ensure DOCX files are properly loaded.");
        }

        try
        {
            // Process the document structure
            var documentResult = await ProcessDocumentAsync();

            // Create ProseMirror document using prosemirror-dotnet
            var documentNode = _proseMirrorSchema.Node(
                _proseMirrorSchema.TopNodeType,
                null,
                Fragment.From(documentResult.Content)
            );

            return documentNode;
        }
        catch (Exception ex)
        {
            if (Debug)
            {
                Console.WriteLine($"Error creating ProseMirror document: {ex.Message}");
            }
            throw new InvalidOperationException("Failed to create ProseMirror document", ex);
        }
    }

    /// <summary>
    /// Process the main document structure
    /// </summary>
    /// <returns>Document processing result</returns>
    private async Task<DocumentProcessingResult> ProcessDocumentAsync()
    {
        if (Debug)
        {
            Console.WriteLine("Starting document processing...");
            Console.WriteLine($"InitialJson structure: {InitialJson}");
        }

        // Newtonsoft.Json XML conversion creates a different structure
        // Look for the document root - it should be something like "w:document"
        JObject documentElement = null;

        // Try to find the document element
        foreach (var property in InitialJson.Properties())
        {
            if (property.Name.Contains("document"))
            {
                documentElement = property.Value as JObject;
                if (Debug)
                {
                    Console.WriteLine($"Found document element: {property.Name}");
                }
                break;
            }
        }

        if (documentElement == null)
        {
            if (Debug)
            {
                Console.WriteLine("Available root elements:");
                foreach (var prop in InitialJson.Properties())
                {
                    Console.WriteLine($"  - {prop.Name}");
                }
            }
            throw new InvalidOperationException("Document root element not found");
        }

        // Look for w:body element within the document
        JObject bodyElement = null;

        foreach (var property in documentElement.Properties())
        {
            if (property.Name.Contains("body"))
            {
                bodyElement = property.Value as JObject;
                if (Debug)
                {
                    Console.WriteLine($"Found body element: {property.Name}");
                }
                break;
            }
        }

        if (bodyElement == null)
        {
            if (Debug)
            {
                Console.WriteLine("Available document elements:");
                foreach (var prop in documentElement.Properties())
                {
                    Console.WriteLine($"  - {prop.Name}");
                }
            }
            throw new InvalidOperationException("Document body not found");
        }

        var content = new List<Node>();

        // Process body content
        // Newtonsoft.Json creates different structures for different content types
        foreach (var bodyProperty in bodyElement.Properties())
        {
            if (Debug)
            {
                Console.WriteLine($"Processing body property: {bodyProperty.Name}");
            }

            // Handle different types of content (paragraphs, tables, etc.)
            if (bodyProperty.Name.Contains("p")) // Paragraphs
            {
                var paragraphData = bodyProperty.Value;
                if (paragraphData is JArray paragraphArray)
                {
                    // Multiple paragraphs
                    foreach (var para in paragraphArray)
                    {
                        var paraObject = para as JObject;
                        if (paraObject != null)
                        {
                            var processedNode = await ProcessParagraphFromNewtonsoftAsync(paraObject);
                            if (processedNode != null)
                            {
                                content.Add(processedNode);
                            }
                        }
                    }
                }
                else if (paragraphData is JObject singleParagraph)
                {
                    // Single paragraph
                    var processedNode = await ProcessParagraphFromNewtonsoftAsync(singleParagraph);
                    if (processedNode != null)
                    {
                        content.Add(processedNode);
                    }
                }
            }
        }

        if (Debug)
        {
            Console.WriteLine($"Processed {content.Count} content elements");
        }

        return new DocumentProcessingResult
        {
            Content = content,
            PageStyles = ExtractPageStyles(),
            Numbering = ExtractNumbering(),
            Comments = ExtractComments()
        };
    }

    /// <summary>
    /// Process paragraph from Newtonsoft.Json XML structure
    /// </summary>
    /// <param name="paragraphObject">Paragraph JSON object</param>
    /// <returns>ProseMirror paragraph node</returns>
    private async Task<Node> ProcessParagraphFromNewtonsoftAsync(JObject paragraphObject)
    {
        var content = new List<Node>();

        if (Debug)
        {
            Console.WriteLine($"Processing paragraph with properties: {string.Join(", ", paragraphObject.Properties().Select(p => p.Name))}");
        }

        // Look for runs (w:r elements) within the paragraph
        foreach (var property in paragraphObject.Properties())
        {
            if (property.Name.Contains("r") && !property.Name.Contains("Pr")) // Runs but not properties
            {
                var runData = property.Value;
                if (runData is JArray runArray)
                {
                    // Multiple runs
                    foreach (var run in runArray)
                    {
                        var runObject = run as JObject;
                        if (runObject != null)
                        {
                            var textNode = await ProcessRunFromNewtonsoftAsync(runObject);
                            if (textNode != null)
                            {
                                content.Add(textNode);
                            }
                        }
                    }
                }
                else if (runData is JObject singleRun)
                {
                    // Single run
                    var textNode = await ProcessRunFromNewtonsoftAsync(singleRun);
                    if (textNode != null)
                    {
                        content.Add(textNode);
                    }
                }
            }
        }

        // If paragraph is empty, add a single space text node (empty text nodes are not allowed)
        if (!content.Any())
        {
            content.Add(_proseMirrorSchema.Text(" "));
        }

        var paragraphType = _proseMirrorSchema.Nodes["paragraph"];
        return paragraphType.Create(null, Fragment.From(content));
    }

    /// <summary>
    /// Process run from Newtonsoft.Json XML structure
    /// </summary>
    /// <param name="runObject">Run JSON object</param>
    /// <returns>ProseMirror text node</returns>
    private async Task<Node> ProcessRunFromNewtonsoftAsync(JObject runObject)
    {
        var textContent = "";

        // Look for text elements (w:t)
        foreach (var property in runObject.Properties())
        {
            if (property.Name.Contains("t")) // Text elements
            {
                var textData = property.Value;
                if (textData is JObject textObject)
                {
                    // Text with attributes
                    if (textObject["#text"] != null)
                    {
                        textContent += textObject["#text"]?.ToString() ?? "";
                    }
                }
                else if (textData is JValue textValue)
                {
                    // Direct text content
                    textContent += textValue.ToString();
                }
            }
        }

        // Only create text node if we have actual content (not just whitespace)
        if (!string.IsNullOrWhiteSpace(textContent))
        {
            return _proseMirrorSchema.Text(textContent);
        }

        // For whitespace-only content, preserve it if it's not empty
        if (!string.IsNullOrEmpty(textContent))
        {
            return _proseMirrorSchema.Text(textContent);
        }

        // Return null instead of empty text node - empty text nodes are not allowed
        return null;
    }

    /// <summary>
    /// Find element recursively by name (with or without namespace)
    /// </summary>
    /// <param name="parent">Parent element</param>
    /// <param name="elementName">Element name to find (e.g., "w:body" or "body")</param>
    /// <returns>Found element or null</returns>
    private JToken FindElementRecursively(JObject parent, string elementName)
    {
        // For Newtonsoft.Json structure, we search through properties
        foreach (var property in parent.Properties())
        {
            if (property.Name == elementName ||
                (elementName == "body" && property.Name.Contains("body")) ||
                (elementName == "w:body" && property.Name.Contains("body")))
            {
                return property.Value;
            }

            // Recursively search in child objects
            if (property.Value is JObject childObj)
            {
                var found = FindElementRecursively(childObj, elementName);
                if (found != null)
                    return found;
            }
            else if (property.Value is JArray childArray)
            {
                foreach (var item in childArray)
                {
                    if (item is JObject itemObj)
                    {
                        var found = FindElementRecursively(itemObj, elementName);
                        if (found != null)
                            return found;
                    }
                }
            }
        }

        return null;
    }

    /// <summary>
    /// Log element structure for debugging
    /// </summary>
    /// <param name="element">Element to log</param>
    /// <param name="indent">Indentation level</param>
    private void LogElementStructure(JObject element, int indent)
    {
        var indentStr = new string(' ', indent * 2);

        foreach (var property in element.Properties())
        {
            Console.WriteLine($"{indentStr}{property.Name}: {property.Value.Type}");

            if (property.Value is JObject childObj && indent < 3) // Limit recursion depth
            {
                LogElementStructure(childObj, indent + 1);
            }
            else if (property.Value is JArray childArray && indent < 3)
            {
                Console.WriteLine($"{indentStr}  Array with {childArray.Count} items");
            }
        }
    }

    // Legacy ProcessElementAsync method removed - replaced with Newtonsoft.Json-based processing

    // Legacy ProcessParagraphAsync method removed - replaced with ProcessParagraphFromNewtonsoftAsync

    // Legacy ProcessTextAsync method removed - replaced with ProcessRunFromNewtonsoftAsync

    // Legacy table processing methods removed - will be reimplemented with Newtonsoft.Json structure

    // Legacy CreateLineBreakNode and ProcessGenericElementAsync methods removed

    #endregion

    #region Helper Methods

    // Legacy helper methods removed - these used JsonObject structure from System.Text.Json.Nodes
    // which is incompatible with the Newtonsoft.Json JObject structure we now use

    /// <summary>
    /// Extract page styles from document
    /// </summary>
    /// <returns>Page styles object</returns>
    private object ExtractPageStyles()
    {
        // Extract page size, margins, orientation from sectPr
        // Implementation needed
        return new { };
    }

    /// <summary>
    /// Extract numbering information
    /// </summary>
    /// <returns>Numbering object</returns>
    private object ExtractNumbering()
    {
        // Extract list definitions from numbering.xml
        // Implementation needed
        return new { };
    }

    /// <summary>
    /// Extract comments from document
    /// </summary>
    /// <returns>Comments list</returns>
    private List<object> ExtractComments()
    {
        // Extract comments from comments.xml
        // Implementation needed
        return new List<object>();
    }

    /// <summary>
    /// Create default ProseMirror schema
    /// This should match the schema used in the JavaScript version
    /// </summary>
    /// <returns>Default schema</returns>
    private Schema CreateDefaultSchema()
    {
        // This is a placeholder - you'll need to create a schema that matches
        // the JavaScript ProseMirror schema used in SuperDoc
        // Refer to the schema definition in the JavaScript codebase

        var spec = new SchemaSpec
        {
            Nodes = new OrderedDictionary<string, NodeSpec>
            {
                ["doc"] = new NodeSpec { Content = "block+" },
                ["paragraph"] = new NodeSpec { Content = "inline*", Group = "block" },
                ["text"] = new NodeSpec { Group = "inline" },
                ["table"] = new NodeSpec { Content = "tableRow+", Group = "block" },
                ["tableRow"] = new NodeSpec { Content = "tableCell+" },
                ["tableCell"] = new NodeSpec { Content = "block+" },
                ["lineBreak"] = new NodeSpec { Group = "inline", Inline = true }
            },
            Marks = new OrderedDictionary<string, MarkSpec>
            {
                ["bold"] = new MarkSpec(),
                ["italic"] = new MarkSpec(),
                ["underline"] = new MarkSpec(),
                ["textStyle"] = new MarkSpec
                {
                    Attrs = new Dictionary<string, AttributeSpec>
                    {
                        ["color"] = new AttributeSpec { Default = null },
                        ["fontSize"] = new AttributeSpec { Default = null },
                        ["fontFamily"] = new AttributeSpec { Default = null }
                    }
                }
            }
        };

        return new Schema(spec);
    }

    #endregion

    #region Export Methods (Future Implementation)

    /// <summary>
    /// Convert ProseMirror document back to DOCX XML
    /// This is the reverse operation - for future implementation
    /// </summary>
    /// <param name="proseMirrorDoc">ProseMirror document</param>
    /// <returns>DOCX XML content</returns>
    public async Task<Dictionary<string, string>> ExportToDocxXmlAsync(Node proseMirrorDoc)
    {
        // Future implementation for export functionality
        // This would convert ProseMirror back to DOCX XML format
        throw new NotImplementedException("Export functionality will be implemented in future versions");
    }

    #endregion
}

#region Supporting Classes

/// <summary>
/// Mark type mapping configuration
/// </summary>
public class MarkTypeMapping
{
    public string Name { get; set; }
    public string Type { get; set; }
    public string Mark { get; set; }
    public string Property { get; set; }
}

/// <summary>
/// Document processing result
/// </summary>
public class DocumentProcessingResult
{
    public List<Node> Content { get; set; } = new();
    public object PageStyles { get; set; }
    public object Numbering { get; set; }
    public List<object> Comments { get; set; } = new();
}

#endregion
