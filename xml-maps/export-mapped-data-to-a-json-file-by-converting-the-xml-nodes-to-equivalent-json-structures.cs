// Title: Recursively convert an XML document to a formatted JSON file with C# JsonNode and XElement
// AI Prompts: Create a C# console program that reads an XML file, recursively maps each element and attribute to JsonNode objects (attributes prefixed with '@' and mixed text stored under '#text'), and writes the indented JSON to a specified output file. | Enhance the conversion method to handle XML namespaces and output property names in camelCase while keeping attribute and array handling intact.
// Common Searches: c# convert xml file to json preserving attributes and arrays | recursive xml to json conversion using System.Text.Json.Nodes | map repeated xml elements to json arrays in c# | include element text and attributes when converting xml to json c# | how to handle xml namespaces during xml to json conversion c#
// Tags: xml to json conversion JsonNode C# | preserve xml attributes as @properties json | xml repeated elements to json arrays | recursive XElement to JsonObject mapping | indented json output C#

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;

// The program loads an XML file, recursively transforms each XElement (including attributes prefixed with '@' and mixed content captured under '#text') into a JsonNode hierarchy, serializes it with indentation, and writes the resulting JSON to an output file.
class Program
{
    static void Main()
    {
        // Paths for input XML and output JSON
        string xmlPath = "input.xml";
        string jsonPath = "output.json";

        // Load the XML document
        XDocument xdoc = XDocument.Load(xmlPath);

        // Convert the root element to a JSON node hierarchy
        JsonNode jsonRoot = ConvertElement(xdoc.Root);

        // Serialize JSON with indentation for readability
        var options = new JsonSerializerOptions { WriteIndented = true };
        string jsonString = jsonRoot.ToJsonString(options);

        // Write the JSON string to a file
        File.WriteAllText(jsonPath, jsonString);
    }

    // Recursively converts an XElement into a JsonNode
    static JsonNode ConvertElement(XElement element)
    {
        // If the element has no children and no attributes, return its text value directly
        if (!element.HasElements && !element.HasAttributes)
        {
            return element.Value;
        }

        var obj = new JsonObject();

        // Include attributes as properties prefixed with '@'
        foreach (var attr in element.Attributes())
        {
            obj[$"@{attr.Name.LocalName}"] = attr.Value;
        }

        // Group child elements by name to handle single vs. multiple occurrences
        var groups = element.Elements().GroupBy(e => e.Name.LocalName);
        foreach (var group in groups)
        {
            if (group.Count() == 1)
            {
                // Single child element becomes a nested object/property
                obj[group.Key] = ConvertElement(group.First());
            }
            else
            {
                // Multiple child elements with the same name become a JSON array
                var array = new JsonArray();
                foreach (var child in group)
                {
                    array.Add(ConvertElement(child));
                }
                obj[group.Key] = array;
            }
        }

        // Preserve element text when mixed with child elements, using "#text" key
        if (!string.IsNullOrWhiteSpace(element.Value) && element.HasElements)
        {
            obj["#text"] = element.Value.Trim();
        }

        return obj;
    }
}
