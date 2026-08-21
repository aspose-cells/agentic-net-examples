// Title: Export Excel XML Map to a .xml File with Aspose.Cells for .NET
// Description: Loads an Excel workbook containing XML maps, verifies map presence, and uses Workbook.ExportXml to write the selected map to a separate .xml file while preserving the original schema. Includes basic error handling for missing files or maps.
// Keywords: Aspose.Cells export XML map | Workbook.ExportXml C# | export XML from Excel | preserve XML schema Aspose | C# export Excel XML map | Aspose.Cells XML map example | save XML map to file
// Common Searches: how to export xml map from excel using aspose.cells | c# workbook.exportxml preserve schema | aspose.cells export xml map to file | export first xml map in workbook aspose | aspose.cells xml map export example
// Developer Intent: Generate an external .xml file from an Excel XML map while keeping the map's schema intact.
// Use Cases: Create a standards‑compliant XML document for data exchange by exporting the workbook's XML map. | Automate validation pipelines that require the original XML schema to remain unchanged. | Integrate Excel‑based data sources into systems that consume XML files, ensuring schema fidelity.
// AI Prompts: Write C# code with Aspose.Cells to export the XML map named "Orders" from "Data.xlsx" to "Orders.xml" preserving the schema. | Show how to list all XML maps in a workbook and export each to separate XML files with error handling for missing maps. | Provide a robust Aspose.Cells example that checks for the input file, verifies XML map existence, and logs detailed errors during export.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook containing XML maps, verifies map presence, and uses Workbook.ExportXml to write the selected map to a separate .xml file while preserving the original schema. Includes basic error handling for missing files or maps.
    public class ExportXmlMapDemo
    {
        public static void Run()
        {
            const string inputFile = "InputWithMap.xlsx";
            const string outputFile = "ExportedData.xml";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
                return;
            }

            try
            {
                // Load the workbook that already contains an XML map
                Workbook workbook = new Workbook(inputFile);

                // Check if any XML maps are present
                if (workbook.Worksheets.XmlMaps.Count > 0)
                {
                    // Retrieve the first XML map (or any specific one you need)
                    XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

                    // Export the XML data linked to this map, preserving the original schema
                    workbook.ExportXml(xmlMap.Name, outputFile);

                    Console.WriteLine($"XML map exported successfully to \"{outputFile}\"");
                }
                else
                {
                    Console.WriteLine("No XML map found in the workbook.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    internal class Program
    {
        private static void Main(string[] args)
        {
            ExportXmlMapDemo.Run();
        }
    }
}
