// Title: Convert XML to formatted JSON file in C# using XmlDocument and System.Text.Json with attribute and array support
// AI Prompts: Write a C# console application that loads an XML file, walks the XmlDocument tree, converts each node into a JsonNode hierarchy, and writes the result to a pretty‑printed JSON file. | Create a ConvertXmlNodeToJsonNode method that maps XML attributes to keys prefixed with '@', element text to a '#text' property, and groups elements with identical names into JsonArray objects using System.Text.Json.
// Common Searches: c# convert xml to json while keeping attributes and element arrays | sample code for mapping XmlNode hierarchy to JsonNode in .NET | generate indented json output from xml document using System.Text.Json
// Tags: xml to json conversion with System.Text.Json | attribute preservation in xml to json mapping | duplicate element handling as json arrays | recursive xml node processing in C# | pretty printed json file generation in .NET

using System;
using System.IO;
using System.Xml;
using System.Text.Json;
using System.Text.Json.Nodes;

// The example loads an XML file with XmlDocument, verifies its structure, and recursively transforms the root XmlNode into a JsonNode tree. Attributes are stored with an '@' prefix, text content uses a '#text' key, and repeated element names are aggregated into JsonArray objects. The resulting JsonNode is serialized with indentation via System.Text.Json and saved to a specified JSON file, with comprehensive error handling for missing files and malformed XML.
class Program
{
    static void Main()
    {
        try
        {
            // Paths for input XML and output JSON
            string xmlPath = "input.xml";
            string jsonPath = "output.json";

            // Verify that the XML file exists
            if (!File.Exists(xmlPath))
            {
                Console.WriteLine($"Error: The file '{xmlPath}' was not found.");
                return;
            }

            // Load the XML document
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlPath);

            // Ensure the document has a root element
            if (xmlDoc.DocumentElement == null)
            {
                Console.WriteLine("Error: The XML document does not contain a root element.");
                return;
            }

            // Convert the root XML node to a JsonNode hierarchy
            JsonNode jsonRoot = ConvertXmlNodeToJsonNode(xmlDoc.DocumentElement);

            // Serialize the JsonNode to a formatted JSON string
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string jsonString = jsonRoot.ToJsonString(jsonOptions);

            // Write the JSON string to the output file
            File.WriteAllText(jsonPath, jsonString);

            Console.WriteLine($"XML has been successfully converted to JSON and saved to '{jsonPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    static JsonNode ConvertXmlNodeToJsonNode(XmlNode xmlNode)
    {
        // Guard against null nodes
        if (xmlNode == null)
            return null;

        // Create a JSON object to represent the current XML element
        JsonObject jsonObj = new JsonObject();

        // Process attributes (if any)
        if (xmlNode.Attributes != null && xmlNode.Attributes.Count > 0)
        {
            foreach (XmlAttribute attr in xmlNode.Attributes)
            {
                // Prefix attribute names with '@' to distinguish them from child elements
                jsonObj[$"@{attr.Name}"] = attr.Value;
            }
        }

        // Process child nodes
        foreach (XmlNode child in xmlNode.ChildNodes)
        {
            switch (child.NodeType)
            {
                case XmlNodeType.Element:
                    // Recursively convert child elements
                    JsonNode childJson = ConvertXmlNodeToJsonNode(child);
                    if (childJson == null)
                        continue;

                    // Handle multiple elements with the same name
                    if (jsonObj.ContainsKey(child.Name))
                    {
                        JsonNode existing = jsonObj[child.Name];
                        if (existing is JsonArray array)
                        {
                            // Add the new element to the existing array
                            array.Add(childJson);
                        }
                        else
                        {
                            // Replace the single object with an array containing both values
                            JsonArray newArray = new JsonArray
                            {
                                // Clone the existing node via serialization to avoid parent conflicts
                                JsonNode.Parse(existing.ToJsonString()),
                                childJson
                            };
                            jsonObj[child.Name] = newArray;
                        }
                    }
                    else
                    {
                        jsonObj[child.Name] = childJson;
                    }
                    break;

                case XmlNodeType.Text:
                case XmlNodeType.CDATA:
                    // Text content of the element
                    string text = child.Value?.Trim();
                    if (!string.IsNullOrEmpty(text))
                    {
                        // Use '#text' as the property name for element text
                        jsonObj["#text"] = text;
                    }
                    break;

                // Ignore comments, processing instructions, etc.
                default:
                    break;
            }
        }

        // If the element has no attributes or child elements, return its text directly
        if (jsonObj.Count == 0 && !string.IsNullOrEmpty(xmlNode.InnerText))
        {
            return JsonValue.Create(xmlNode.InnerText);
        }

        return jsonObj;
    }
}
