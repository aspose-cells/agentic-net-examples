using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using System.Xml.Linq;

namespace AsposeCellsJsonExport
{
    class Program
    {
        static void Main()
        {
            // Load the workbook that contains the XML map
            Workbook workbook = new Workbook("input.xlsx");

            // Ensure there is at least one XML map defined
            if (workbook.Worksheets.XmlMaps.Count == 0)
            {
                Console.WriteLine("No XML map found in the workbook.");
                return;
            }

            // Get the first XML map (you can select a specific one by name if needed)
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

            // Export the XML data linked to the map into a memory stream
            using (MemoryStream xmlStream = new MemoryStream())
            {
                workbook.ExportXml(xmlMap.Name, xmlStream);
                xmlStream.Position = 0; // Reset stream position for reading

                // Load the exported XML into an XDocument
                XDocument xDoc = XDocument.Load(xmlStream);

                // Convert the XDocument to JSON using System.Text.Json
                string json = JsonSerializer.Serialize(xDoc, new JsonSerializerOptions { WriteIndented = true });

                // Define the output JSON file path
                string jsonPath = "mappedData.json";

                // Write the JSON string to the file
                File.WriteAllText(jsonPath, json);

                Console.WriteLine($"Mapped data exported to JSON successfully at '{jsonPath}'.");
            }
        }
    }
}