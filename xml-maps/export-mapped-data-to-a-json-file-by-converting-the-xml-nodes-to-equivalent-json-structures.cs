// Title: How to export an Excel XML map to a JSON file with Aspose.Cells for .NET (C#)
// Common Searches: c# export xml map from excel workbook to json using aspose.cells | aspose.cells ExportXml method example for json output | convert excel xml map to json file with aspose.cells .net library | sample code for exporting xml map data to json in c# | aspose.cells exportxml generate json from mapped data
// Tags: Aspose.Cells ExportXml JSON output | Excel XML map to JSON conversion C# | Export XML map from workbook using Aspose.Cells | XElement hierarchical to JSON serialization .NET | MemoryStream XML export Aspose.Cells example

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsXmlToJsonExport
{
    // The C# program loads an Excel workbook, checks for defined XML maps, uses Aspose.Cells' ExportXml method to write the first XML map to a MemoryStream, parses the XML with XDocument, recursively converts elements and attributes into a hierarchical object, serializes the object to indented JSON, and saves the result as a JSON file.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths – adjust as needed
            const string inputWorkbookPath = "MappedData.xlsx";
            const string outputJsonPath = "MappedData.json";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputWorkbookPath))
                {
                    Console.WriteLine($"Error: Input workbook \"{inputWorkbookPath}\" not found.");
                    return;
                }

                // Load the workbook
                using (Workbook workbook = new Workbook(inputWorkbookPath))
                {
                    // Ensure there is at least one XML map defined
                    if (workbook.Worksheets.XmlMaps.Count == 0)
                    {
                        Console.WriteLine("Error: No XML maps are defined in the workbook.");
                        return;
                    }

                    // Use the first XML map (you can change this logic to select a specific map)
                    XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];
                    Console.WriteLine($"Using XML map: {xmlMap.Name}");

                    // Export the XML data linked by the map into a memory stream
                    using (MemoryStream xmlStream = new MemoryStream())
                    {
                        workbook.ExportXml(xmlMap.Name, xmlStream);
                        xmlStream.Position = 0; // Reset stream position for reading

                        // Load the exported XML into XDocument
                        XDocument xDoc = XDocument.Load(xmlStream);
                        Console.WriteLine("XML export successful.");

                        // Convert the XDocument to a JSON structure
                        object jsonObject = XElementToObject(xDoc.Root);
                        var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                        string jsonString = JsonSerializer.Serialize(jsonObject, jsonOptions);

                        // Write JSON to file
                        File.WriteAllText(outputJsonPath, jsonString);
                        Console.WriteLine($"JSON file created at \"{outputJsonPath}\".");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        private static object XElementToObject(XElement element)
        {
            // If the element has no child elements, return its value (or attributes if any)
            if (!element.HasElements)
            {
                // Include attributes as a dictionary if they exist
                if (element.HasAttributes)
                {
                    var dict = new Dictionary<string, object>();
                    foreach (var attr in element.Attributes())
                    {
                        dict[$"@{attr.Name.LocalName}"] = attr.Value;
                    }
                    dict["#text"] = element.Value;
                    return dict;
                }
                return element.Value;
            }

            // Element has child elements – group them
            var childGroups = element.Elements()
                                     .GroupBy(e => e.Name.LocalName)
                                     .ToDictionary(g => g.Key, g => g.ToList());

            var result = new Dictionary<string, object>();

            // Process attributes
            foreach (var attr in element.Attributes())
            {
                result[$"@{attr.Name.LocalName}"] = attr.Value;
            }

            // Process child groups
            foreach (var kvp in childGroups)
            {
                string childName = kvp.Key;
                List<XElement> childElements = kvp.Value;

                if (childElements.Count == 1)
                {
                    // Single child – store as object
                    result[childName] = XElementToObject(childElements[0]);
                }
                else
                {
                    // Multiple children with same name – store as array
                    var list = new List<object>();
                    foreach (var child in childElements)
                    {
                        list.Add(XElementToObject(child));
                    }
                    result[childName] = list;
                }
            }

            return result;
        }
    }
}
