using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using YDotNet.Document;
using YDotNet.Document.Types.Maps;
using YDotNet.Document.Transactions;
using YDotNet.Document.Types.XmlFragments;
using YDotNet.Document.Types.XmlElements;
using YDotNet.Document.Types.XmlTexts;
using YDotNet.Document.Cells;
using StepWise.Prose.Model;
using StepWise.Prose.Collections;
using SuperDocPoc.Collaboration;

namespace SuperDocPoc.Converter;

public class DocxToYdocService
{
    public static async Task<Doc> ConvertDocxToYDocAsync(Stream docxFileStream)
    {
        try
        {
            var conversionService = new DocxToProseMirrorService();
            var conversionResult = await conversionService.ConvertDocxToProseMirrorAsync(docxFileStream);

            if (!conversionResult.Success)
            {
                throw new InvalidOperationException($"DOCX conversion failed: {conversionResult.Error}");
            }

            var collaborationData = conversionService.PrepareCollaborationData(conversionResult);

            var ydoc = new Doc();

            // CORRECT PATTERN: Store ProseMirror document as XML Fragment (like JavaScript)
            // This matches: const fragment = ydoc.getXmlFragment('supereditor');
            var fragment = ydoc.XmlFragment("supereditor");
            var mediaMap = ydoc.Map("media");
            var metaMap = ydoc.Map("meta");

            // Create a transaction for all operations
            using var transaction = ydoc.WriteTransaction();

            // Implement proper ProseMirror Node to Y.js XML Fragment conversion
            // This is the C# equivalent of JavaScript's prosemirrorToYDoc(editor.state.doc, 'supereditor')
            ConvertProseMirrorToYjsFragment(conversionResult.ProseMirrorDocument, fragment, transaction);

            // Store media files in media map (matches JavaScript pattern)
            foreach (var media in collaborationData.MediaMap)
            {
                // Handle different media value types
                var mediaValue = media.Value switch
                {
                    string stringValue => stringValue,
                    byte[] byteArrayValue => Convert.ToBase64String(byteArrayValue),
                    _ => media.Value?.ToString() ?? ""
                };
                mediaMap.Insert(transaction, media.Key, Input.String(mediaValue));
            }

            // Store metadata in meta map (matches JavaScript pattern)            
            foreach (var meta in collaborationData.MetaData)
            {
                var metaValue = meta.Value?.ToString() ?? "";
                metaMap.Insert(transaction, meta.Key, Input.String(metaValue));
            }

            // Store DOCX data in meta map (following JavaScript collaboration-helpers.js pattern)
            // This matches: metaMap.set('docx', editor.options.content);
            // where editor.options.content is an array of { name, content } objects
            var docxArray = new List<object>();
            foreach (var xmlFile in collaborationData.XmlFiles)
            {
                docxArray.Add(new { name = xmlFile.Key, content = xmlFile.Value });
            }
            var docxJson = JsonConvert.SerializeObject(docxArray);
            metaMap.Insert(transaction, "docx", Input.String(docxJson));

            // Transaction will be committed automatically when disposed

            return ydoc;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to convert DOCX to Y.js document: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Convert ProseMirror Node to Y.js XML Fragment - C# equivalent of prosemirrorToYDoc()
    /// This converts the ProseMirror document structure to Y.js collaborative format
    /// </summary>
    /// <param name="proseMirrorNode">The ProseMirror document node</param>
    /// <param name="xmlFragment">The Y.js XML fragment to populate</param>
    /// <param name="transaction">The transaction for Y.js operations</param>
    private static void ConvertProseMirrorToYjsFragment(Node proseMirrorNode, XmlFragment xmlFragment, Transaction transaction)
    {
        if (proseMirrorNode == null)
        {
            return;
        }

        // Convert the ProseMirror node and its content to Y.js XML structure
        ConvertNodeToYjsXml(proseMirrorNode, xmlFragment, null, transaction);
    }

    /// <summary>
    /// Recursively convert ProseMirror nodes to Y.js XML elements
    /// </summary>
    /// <param name="node">ProseMirror node to convert</param>
    /// <param name="parentFragment">Parent Y.js XML fragment</param>
    /// <param name="parentElement">Parent Y.js XML element (if nested)</param>
    /// <param name="transaction">The transaction for Y.js operations</param>
    private static void ConvertNodeToYjsXml(Node node, XmlFragment parentFragment, XmlElement parentElement, Transaction transaction)
    {
        if (node.IsText)
        {
            // Handle text nodes
            var textContent = node.Text ?? "";

            if (parentElement != null)
            {
                // Insert text into XML element
                var childCount = parentElement.ChildLength(transaction);
                var xmlText = parentElement.InsertText(transaction, childCount);
                xmlText.Insert(transaction, 0, textContent);
            }
            else
            {
                // Insert text directly into fragment
                var childCount = parentFragment.ChildLength(transaction);
                var xmlText = parentFragment.InsertText(transaction, childCount);
                xmlText.Insert(transaction, 0, textContent);
            }
            return;
        }

        // Create XML element for this node
        var nodeName = node.Type.Name;
        XmlElement xmlElement;

        if (parentElement != null)
        {
            var childCount = parentElement.ChildLength(transaction);
            xmlElement = parentElement.InsertElement(transaction, childCount, nodeName);
        }
        else
        {
            var childCount = parentFragment.ChildLength(transaction);
            xmlElement = parentFragment.InsertElement(transaction, childCount, nodeName);
        }

        // Add node attributes as XML attributes
        if (node.Attrs != null)
        {
            var attrsJson = JObject.FromObject(node.Attrs);
            foreach (var attr in attrsJson.Properties())
            {
                if (attr.Value != null && !attr.Value.Type.Equals(JTokenType.Null))
                {
                    xmlElement.InsertAttribute(transaction, attr.Name, attr.Value.ToString());
                }
            }
        }

        // Add node marks as attributes (equivalent to ProseMirror marks)
        if (node.Marks != null && node.Marks.Count > 0)
        {
            var marksData = new List<object>();
            foreach (var mark in node.Marks)
            {
                var markData = new
                {
                    type = mark.Type.Name,
                    attrs = mark.Attrs
                };
                marksData.Add(markData);
            }

            var marksJson = JsonConvert.SerializeObject(marksData);
            xmlElement.InsertAttribute(transaction, "marks", marksJson);
        }

        // Recursively process child nodes
        if (node.Content != null && node.Content.ChildCount > 0)
        {
            for (int i = 0; i < node.Content.ChildCount; i++)
            {
                var childNode = node.Content.Child(i);
                ConvertNodeToYjsXml(childNode, parentFragment, xmlElement, transaction);
            }
        }
    }
}
