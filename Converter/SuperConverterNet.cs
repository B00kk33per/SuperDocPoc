using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;
using StepWise.Prose.Collections;
using StepWise.Prose.Model;
using System.Xml;

namespace SuperDocPoc.Converter;

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
    public Dictionary<string, JsonObject> ConvertedXml { get; private set; } = new();
    public Dictionary<string, object> Media { get; private set; } = new();
    public Dictionary<string, object> Fonts { get; private set; } = new();
    public Dictionary<string, object> AddedMedia { get; private set; } = new();
    public List<object> Comments { get; private set; } = new();
    public JsonObject InitialJson { get; private set; }
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
                if (file.Key == "word/document.xml" && ConvertedXml[file.Key]["elements"] is JsonArray elements && elements.Count > 0)
                {
                    var docElement = elements[0] as JsonObject;
                    // Store document attributes if needed
                }
            }
            catch (Exception ex)
            {
                if (Debug)
                {
                    Console.WriteLine($"Error parsing {file.Key}: {ex.Message}");
                }
                throw new InvalidOperationException($"Failed to parse XML file {file.Key}", ex);
            }
        }

        InitialJson = ConvertedXml.ContainsKey("word/document.xml") ? ConvertedXml["word/document.xml"] : null;

        // Extract XML declaration if available
        if (InitialJson?["declaration"] != null)
        {
            // Store declaration for later use in export
        }
    }

    /// <summary>
    /// Convert XML string to JSON structure (equivalent to xml-js)
    /// </summary>
    /// <param name="xml">XML content</param>
    /// <returns>JSON representation of XML</returns>
    private JsonObject ParseXmlToJson(string xml)
    {
        try
        {
            var xDoc = XDocument.Parse(xml);
            return ConvertXElementToJson(xDoc.Root);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to parse XML: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Convert XElement to JSON format similar to xml-js output
    /// </summary>
    /// <param name="element">XElement to convert</param>
    /// <returns>JSON object</returns>
    private JsonObject ConvertXElementToJson(XElement element)
    {
        var result = new JsonObject();

        if (element.Document?.Declaration != null)
        {
            result["declaration"] = new JsonObject
            {
                ["attributes"] = new JsonObject
                {
                    ["version"] = element.Document.Declaration.Version,
                    ["encoding"] = element.Document.Declaration.Encoding,
                    ["standalone"] = element.Document.Declaration.Standalone
                }
            };
        }

        var elements = new JsonArray();
        elements.Add(ConvertElementToJson(element));
        result["elements"] = elements;

        return result;
    }

    /// <summary>
    /// Convert individual XElement to JSON
    /// </summary>
    /// <param name="element">XElement to convert</param>
    /// <returns>JSON object</returns>
    private JsonObject ConvertElementToJson(XElement element)
    {
        var jsonElement = new JsonObject
        {
            ["type"] = "element",
            ["name"] = element.Name.LocalName
        };

        // Add attributes
        if (element.Attributes().Any())
        {
            var attributes = new JsonObject();
            foreach (var attr in element.Attributes())
            {
                attributes[attr.Name.LocalName] = attr.Value;
            }
            jsonElement["attributes"] = attributes;
        }

        // Add child elements and text
        var childElements = new JsonArray();

        foreach (var child in element.Nodes())
        {
            if (child is XElement childElement)
            {
                childElements.Add(ConvertElementToJson(childElement));
            }
            else if (child is XText textNode && !string.IsNullOrWhiteSpace(textNode.Value))
            {
                childElements.Add(new JsonObject
                {
                    ["type"] = "text",
                    ["text"] = textNode.Value
                });
            }
        }

        if (childElements.Any())
        {
            jsonElement["elements"] = childElements;
        }

        return jsonElement;
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
        var documentElements = InitialJson["elements"] as JsonArray;
        var docElement = documentElements?[0] as JsonObject;
        var bodyElements = docElement?["elements"] as JsonArray;
        var bodyElement = bodyElements?.FirstOrDefault(e => (e as JsonObject)?["name"]?.ToString() == "w:body") as JsonObject;

        if (bodyElement == null)
        {
            throw new InvalidOperationException("Document body not found");
        }

        var content = new List<Node>();
        var bodyChildElements = bodyElement["elements"] as JsonArray;

        if (bodyChildElements != null)
        {
            foreach (var element in bodyChildElements.OfType<JsonObject>())
            {
                var processedNode = await ProcessElementAsync(element);
                if (processedNode != null)
                {
                    content.Add(processedNode);
                }
            }
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
    /// Process individual XML element and convert to ProseMirror node
    /// </summary>
    /// <param name="element">XML element as JSON</param>
    /// <returns>ProseMirror node or null</returns>
    private async Task<Node> ProcessElementAsync(JsonObject element)
    {
        var elementName = element["name"]?.ToString();

        if (string.IsNullOrEmpty(elementName) || !AllowedElements.ContainsKey(elementName))
        {
            if (Debug)
            {
                Console.WriteLine($"Skipping unsupported element: {elementName}");
            }
            return null;
        }

        var proseMirrorType = AllowedElements[elementName];

        return proseMirrorType switch
        {
            "paragraph" => await ProcessParagraphAsync(element),
            "text" => await ProcessTextAsync(element),
            "table" => await ProcessTableAsync(element),
            "tableRow" => await ProcessTableRowAsync(element),
            "tableCell" => await ProcessTableCellAsync(element),
            "lineBreak" => CreateLineBreakNode(),
            _ => await ProcessGenericElementAsync(element, proseMirrorType)
        };
    }

    /// <summary>
    /// Process paragraph element
    /// </summary>
    /// <param name="element">Paragraph XML element</param>
    /// <returns>ProseMirror paragraph node</returns>
    private async Task<Node> ProcessParagraphAsync(JsonObject element)
    {
        var content = new List<Node>();
        var attrs = ExtractParagraphAttributes(element);
        var marks = ExtractMarksFromProperties(element);

        var childElements = element["elements"] as JsonArray;
        if (childElements != null)
        {
            foreach (var child in childElements.OfType<JsonObject>())
            {
                var childNode = await ProcessElementAsync(child);
                if (childNode != null)
                {
                    content.Add(childNode);
                }
            }
        }

        // If paragraph is empty, add empty text node
        if (!content.Any())
        {
            content.Add(_proseMirrorSchema.Text(""));
        }

        var paragraphType = _proseMirrorSchema.Nodes["paragraph"];
        return paragraphType.Create(attrs, Fragment.From(content), marks);
    }

    /// <summary>
    /// Process text element (w:t or w:delText)
    /// </summary>
    /// <param name="element">Text XML element</param>
    /// <returns>ProseMirror text node</returns>
    private async Task<Node> ProcessTextAsync(JsonObject element)
    {
        var textContent = ExtractTextContent(element);
        var marks = ExtractMarksFromRunProperties(element);

        return _proseMirrorSchema.Text(textContent, marks);
    }

    /// <summary>
    /// Process table element
    /// </summary>
    /// <param name="element">Table XML element</param>
    /// <returns>ProseMirror table node</returns>
    private async Task<Node> ProcessTableAsync(JsonObject element)
    {
        var content = new List<Node>();
        var attrs = ExtractTableAttributes(element);

        var childElements = element["elements"] as JsonArray;
        if (childElements != null)
        {
            foreach (var child in childElements.OfType<JsonObject>())
            {
                if (child["name"]?.ToString() == "w:tr")
                {
                    var rowNode = await ProcessTableRowAsync(child);
                    if (rowNode != null)
                    {
                        content.Add(rowNode);
                    }
                }
            }
        }

        var tableType = _proseMirrorSchema.Nodes["table"];
        return tableType.Create(attrs, Fragment.From(content));
    }

    /// <summary>
    /// Process table row element
    /// </summary>
    /// <param name="element">Table row XML element</param>
    /// <returns>ProseMirror table row node</returns>
    private async Task<Node> ProcessTableRowAsync(JsonObject element)
    {
        var content = new List<Node>();
        var attrs = ExtractTableRowAttributes(element);

        var childElements = element["elements"] as JsonArray;
        if (childElements != null)
        {
            foreach (var child in childElements.OfType<JsonObject>())
            {
                if (child["name"]?.ToString() == "w:tc")
                {
                    var cellNode = await ProcessTableCellAsync(child);
                    if (cellNode != null)
                    {
                        content.Add(cellNode);
                    }
                }
            }
        }

        var tableRowType = _proseMirrorSchema.Nodes["tableRow"];
        return tableRowType.Create(attrs, Fragment.From(content));
    }

    /// <summary>
    /// Process table cell element
    /// </summary>
    /// <param name="element">Table cell XML element</param>
    /// <returns>ProseMirror table cell node</returns>
    private async Task<Node> ProcessTableCellAsync(JsonObject element)
    {
        var content = new List<Node>();
        var attrs = ExtractTableCellAttributes(element);

        var childElements = element["elements"] as JsonArray;
        if (childElements != null)
        {
            foreach (var child in childElements.OfType<JsonObject>())
            {
                var childNode = await ProcessElementAsync(child);
                if (childNode != null)
                {
                    content.Add(childNode);
                }
            }
        }

        // Table cells must have at least one paragraph
        if (!content.Any())
        {
            var paragraphType = _proseMirrorSchema.Nodes["paragraph"];
            content.Add(paragraphType.Create(null, Fragment.From(_proseMirrorSchema.Text(""))));
        }

        var tableCellType = _proseMirrorSchema.Nodes["tableCell"];
        return tableCellType.Create(attrs, Fragment.From(content));
    }

    /// <summary>
    /// Create line break node
    /// </summary>
    /// <returns>ProseMirror line break node</returns>
    private Node CreateLineBreakNode()
    {
        var lineBreakType = _proseMirrorSchema.Nodes["lineBreak"];
        return lineBreakType.Create();
    }

    /// <summary>
    /// Process generic element
    /// </summary>
    /// <param name="element">XML element</param>
    /// <param name="proseMirrorType">Target ProseMirror type</param>
    /// <returns>ProseMirror node</returns>
    private async Task<Node> ProcessGenericElementAsync(JsonObject element, string proseMirrorType)
    {
        // Handle other element types as needed
        // This is a placeholder for additional element types

        if (Debug)
        {
            Console.WriteLine($"Processing generic element: {proseMirrorType}");
        }

        return null; // Return null for unsupported elements
    }

    #endregion

    #region Helper Methods

    /// <summary>
    /// Extract text content from XML element
    /// </summary>
    /// <param name="element">XML element</param>
    /// <returns>Text content</returns>
    private string ExtractTextContent(JsonObject element)
    {
        var childElements = element["elements"] as JsonArray;
        if (childElements == null) return "";

        var textBuilder = new System.Text.StringBuilder();
        foreach (var child in childElements.OfType<JsonObject>())
        {
            if (child["type"]?.ToString() == "text")
            {
                textBuilder.Append(child["text"]?.ToString() ?? "");
            }
        }

        return textBuilder.ToString();
    }

    /// <summary>
    /// Extract marks from run properties (w:rPr)
    /// </summary>
    /// <param name="element">XML element</param>
    /// <returns>List of ProseMirror marks</returns>
    private List<Mark> ExtractMarksFromRunProperties(JsonObject element)
    {
        var marks = new List<Mark>();

        // Find parent run element and its properties
        // This is a simplified implementation - you'll need to traverse the element tree
        // to find run properties and convert them to ProseMirror marks

        return marks;
    }

    /// <summary>
    /// Extract marks from paragraph properties
    /// </summary>
    /// <param name="element">XML element</param>
    /// <returns>List of ProseMirror marks</returns>
    private List<Mark> ExtractMarksFromProperties(JsonObject element)
    {
        var marks = new List<Mark>();

        // Extract paragraph-level formatting
        // Implementation needed based on specific requirements

        return marks;
    }

    /// <summary>
    /// Extract paragraph attributes
    /// </summary>
    /// <param name="element">Paragraph XML element</param>
    /// <returns>Attributes dictionary</returns>
    private Attrs? ExtractParagraphAttributes(JsonObject element)
    {
        // Extract paragraph properties from w:pPr
        // Implementation needed based on specific requirements

        // For now, return null - you can implement specific attribute extraction here
        return null;
    }

    /// <summary>
    /// Extract table attributes
    /// </summary>
    /// <param name="element">Table XML element</param>
    /// <returns>Attributes dictionary</returns>
    private Attrs? ExtractTableAttributes(JsonObject element)
    {
        // Extract table properties
        // Implementation needed

        return null;
    }

    /// <summary>
    /// Extract table row attributes
    /// </summary>
    /// <param name="element">Table row XML element</param>
    /// <returns>Attributes dictionary</returns>
    private Attrs? ExtractTableRowAttributes(JsonObject element)
    {
        // Extract table row properties
        // Implementation needed

        return null;
    }

    /// <summary>
    /// Extract table cell attributes
    /// </summary>
    /// <param name="element">Table cell XML element</param>
    /// <returns>Attributes dictionary</returns>
    private Attrs? ExtractTableCellAttributes(JsonObject element)
    {
        // Extract table cell properties like colspan, rowspan
        // Implementation needed

        return null;
    }

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
