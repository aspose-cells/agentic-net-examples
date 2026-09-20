// Title: Export a selected XML map from an Excel workbook to an XML file using Aspose.Cells Workbook.ExportXml in C#
// AI Prompts: Generate C# code that opens an .xlsx workbook, verifies that XML maps exist, selects the map at a specified zero‑based index, and calls Workbook.ExportXml to write the map to an .xml file. | Write a C# example that validates a map index against the workbook's XmlMaps collection, retrieves the corresponding map name, and exports that XML map using Aspose.Cells.
// Common Searches: C# Aspose.Cells export XML map by index from Excel workbook | How to use Workbook.ExportXml to save a specific XML map in .NET | Retrieve and export a particular XML map from an .xlsx file using Aspose.Cells | Export selected XML map to file with Aspose.Cells C# example | Validate XML map index before exporting with Aspose.Cells
// Tags: Aspose.Cells specific XML map export | Workbook.ExportXml with map name C# | select xml map using zero-based index Aspose.Cells | check XmlMaps collection before export .NET | write XML map to file Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an Excel workbook, ensures it contains XML maps, validates a zero‑based map index, obtains the map's name, and uses Workbook.ExportXml to save the chosen XML map to an output file, with error handling for missing files or invalid indices.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that may contain XML maps
            Workbook workbook = new Workbook(inputPath);
            dynamic wbDyn = workbook; // Use dynamic to access XmlMaps if available at runtime

            // Index of the map to export (0‑based). Change as needed.
            int mapIndex = 0;

            // Check if XmlMaps property exists and contains maps
            if (wbDyn.XmlMaps == null || wbDyn.XmlMaps.Count == 0)
            {
                Console.WriteLine("No XML maps found in the workbook.");
                return;
            }

            // Validate the requested map index
            if (mapIndex < 0 || mapIndex >= wbDyn.XmlMaps.Count)
            {
                Console.WriteLine($"Map index {mapIndex} is out of range. Available maps: {wbDyn.XmlMaps.Count}");
                return;
            }

            // Retrieve the map name by index
            string mapName = wbDyn.XmlMaps[mapIndex].Name;

            // Export the XML data for the specified map to a file
            string outputPath = "output.xml";
            workbook.ExportXml(outputPath, mapName);
            Console.WriteLine($"XML map exported successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
