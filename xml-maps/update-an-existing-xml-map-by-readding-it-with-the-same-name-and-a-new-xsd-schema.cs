// Title: Replace an existing XML map in an Excel workbook with a new XSD schema using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# replace XML map in an existing workbook with a new XSD schema | remove an XML map and add a new one with the same name using Aspose.Cells .NET | programmatically update the XML map schema of an .xlsx file with Aspose.Cells | export XML data after updating the workbook's XML map in C# | how to re‑add an XML map after deleting it in Aspose.Cells for .NET
// Tags: xml map removal Aspose.Cells C# | add xml map from xsd Aspose.Cells | update workbook xml map schema .NET | export workbook xml using Aspose.Cells | rebind xml map after schema change Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace UpdateXmlMapDemo
{
    // Demonstrates loading a workbook, removing an existing XML map, adding a new XML map from a specified XSD file while preserving the original map name, saving the workbook, and exporting XML data using the updated map.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to required files – adjust these paths as needed.
            string workbookPath = "input.xlsx";
            string newXsdPath = "newSchema.xsd";
            string outputWorkbookPath = "output_updated.xlsx";
            string exportedXmlPath = "exported.xml";

            // Validate input files.
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                return;
            }

            if (!File.Exists(newXsdPath))
            {
                Console.WriteLine($"Error: XSD schema file not found at '{newXsdPath}'.");
                return;
            }

            try
            {
                // Load the existing workbook.
                Workbook workbook = new Workbook(workbookPath);
                Worksheet worksheet = workbook.Worksheets[0]; // Assuming the map is on the first sheet.

                // Access the collection of XML maps.
                XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

                if (xmlMaps.Count == 0)
                {
                    Console.WriteLine("No XML maps exist in the workbook. Adding a new one.");
                }

                // Determine the name of the map we want to replace.
                // For this demo we take the first map's name if any; otherwise we define a name.
                string targetMapName;
                int existingMapIndex = -1;

                if (xmlMaps.Count > 0)
                {
                    // Find the first map (or you could search by a specific name).
                    XmlMap existingMap = xmlMaps[0];
                    targetMapName = existingMap.Name;
                    existingMapIndex = 0;
                }
                else
                {
                    // Define a default name for the new map.
                    targetMapName = "MyXmlMap";
                }

                // If an existing map with the target name is present, remove it.
                if (existingMapIndex >= 0)
                {
                    // XmlMapCollection inherits from CollectionBase<T>, which provides RemoveAt.
                    xmlMaps.RemoveAt(existingMapIndex);
                    Console.WriteLine($"Removed existing XML map '{targetMapName}'.");
                }

                // Add the new XML map using the XSD file.
                // The Add method expects a path or URL to an XSD/XML file.
                int newMapIndex = xmlMaps.Add(newXsdPath);
                XmlMap newMap = xmlMaps[newMapIndex];

                // Set the name of the new map to the original name.
                newMap.Name = targetMapName;
                Console.WriteLine($"Added new XML map with name '{newMap.Name}' using schema '{newXsdPath}'.");

                // OPTIONAL: If you need to bind the map to a range, you can do it here.
                // For example, bind the map to the used range of the worksheet:
                // newMap.DataBinding = new XmlDataBinding(worksheet.Name, worksheet.Cells.MaxDisplayRange.StartRow, worksheet.Cells.MaxDisplayRange.StartColumn);

                // Save the workbook with the updated map.
                workbook.Save(outputWorkbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved with updated XML map to '{outputWorkbookPath}'.");

                // Demonstrate exporting data using the updated map.
                // This will create an XML file based on the new schema.
                workbook.ExportXml(newMap.Name, exportedXmlPath);
                Console.WriteLine($"Exported XML using the updated map to '{exportedXmlPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while updating the XML map:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
