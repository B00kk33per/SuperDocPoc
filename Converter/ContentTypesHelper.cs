using System.Xml.Linq;

namespace SuperDocPoc.Converter;

public static class ContentTypesHelper
{
    /// <summary>
    /// Get content types from [Content_Types].xml
    /// </summary>
    /// <param name="contentTypesXml">The XML content of [Content_Types].xml</param>
    /// <returns>Array of file extensions that are already defined</returns>
    public static string[] GetContentTypesFromXml(string contentTypesXml)
    {
        try
        {
            var xmlDoc = XDocument.Parse(contentTypesXml);
            var namespaceUri = xmlDoc.Root?.Name.Namespace ?? XNamespace.None;

            return xmlDoc.Root?
                .Elements(namespaceUri + "Default")?
                .Select(el => el.Attribute("Extension")?.Value ?? string.Empty)
                .Where(ext => !string.IsNullOrEmpty(ext))
                .ToArray() ?? [];
        }
        catch (Exception)
        {
            // Return empty list if XML parsing fails
            return [];
        }
    }
}
