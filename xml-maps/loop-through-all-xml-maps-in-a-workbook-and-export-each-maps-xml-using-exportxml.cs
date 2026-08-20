// Title: Export Every XML Map in an Excel Workbook to Individual Files with Aspose.Cells for .NET
// Description: Loads a workbook, accesses its XmlMapCollection, iterates through each XmlMap, and calls Workbook.ExportXml to write each map to a distinct .xml file while handling missing files and runtime errors.
// Keywords: Aspose.Cells | C# ExportXml | XML map collection | Excel to XML conversion | batch export XML maps | Workbook.ExportXml | Aspose.Cells .NET
// Common Searches: export all xml maps Aspose.Cells | loop through XmlMapCollection C# | Workbook.ExportXml multiple maps example | save each Excel XML map to separate file | Aspose.Cells export xml maps batch
// Developer Intent: Create separate XML files for every XML map defined in an Excel workbook.
// Use Cases: Provide downstream systems with individual XML files for each data schema embedded in a workbook. | Back up all XML map data before performing bulk edits or migrations. | Automate extraction of XML map contents for a reporting pipeline that consolidates data from multiple maps.
// AI Prompts: Generate C# code that detects duplicate XML map names and appends a numeric suffix before exporting each map with Aspose.Cells ExportXml. | Show how to export all XML maps to a chosen output directory and resolve file‑name conflicts using Aspose.Cells for .NET. | Provide an example that logs the export result of each XML map and captures any ExportXml exceptions for later analysis.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, accesses its XmlMapCollection, iterates through each XmlMap, and calls Workbook.ExportXml to write each map to a distinct .xml file while handling missing files and runtime errors.
    public class ExportAllXmlMapsDemo
    {
        public static void Run()
        {
            string inputPath = "input.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook from an existing Excel file
                Workbook workbook = new Workbook(inputPath);

                // Access the collection of XML maps in the workbook
                XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

                // Iterate through each XML map and export its XML data
                for (int i = 0; i < xmlMaps.Count; i++)
                {
                    XmlMap map = xmlMaps[i];
                    string outputPath = $"{map.Name}.xml";

                    // Export the XML data for the current map
                    workbook.ExportXml(map.Name, outputPath);
                    Console.WriteLine($"Exported XML map '{map.Name}' to '{outputPath}'.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportAllXmlMapsDemo.Run();
        }
    }
}
