// Title: How to remove an XML map and clear its linked cells in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file, locates the XML map called 'MyXmlMap' via the XmlMaps collection, invokes ClearAllLinkedCells, deletes the map, and saves the workbook. | Generate a C# example that uses reflection to access the XmlMaps property in Aspose.Cells, safely clears all cells linked to a specific XML map, removes the map, and handles cases where the library version lacks XML map support. | Create a C# snippet that checks for XML map support, clears linked cells of a named map, removes the map from the workbook, and writes the result to a new file.
// Common Searches: c# aspocells delete specific xml map and clear linked cells | how to clear all linked cells after removing an xml map with Aspose.Cells | using reflection to access XmlMaps collection in Aspose.Cells .NET | aspocells version check for xml map support before deletion | remove xml map named MyXmlMap from Excel workbook using Aspose.Cells
// Tags: aspocells delete xml map c# | aspocells clear linked cells xml map | xml map deletion via reflection aspocells | aspocells xmlmaps collection compatibility | c# xml map removal aspocells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// This C# example loads an Excel workbook, uses reflection to obtain the XmlMaps collection, finds the XML map named 'MyXmlMap', attempts to clear all cells linked to that map with ClearAllLinkedCells, removes the map, and saves the workbook, while handling versions of Aspose.Cells that may not support XML maps.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string xmlMapName = "MyXmlMap";

            // Verify input file existence
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load workbook
            Workbook workbook = new Workbook(inputPath);

            // Use reflection to obtain XmlMaps collection (may not be present in older versions)
            PropertyInfo xmlMapsProp = workbook.GetType().GetProperty("XmlMaps", BindingFlags.Public | BindingFlags.Instance);
            if (xmlMapsProp == null)
            {
                Console.WriteLine("The loaded Aspose.Cells version does not support XML maps.");
                // Save workbook unchanged and exit
                EnsureOutputDirectory(outputPath);
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
                return;
            }

            object xmlMapsObj = xmlMapsProp.GetValue(workbook);
            if (xmlMapsObj == null)
            {
                Console.WriteLine("Failed to retrieve XML maps collection.");
                return;
            }

            dynamic xmlMaps = xmlMapsObj; // Use dynamic for runtime member access
            int mapIndex = -1;

            // Locate the XML map by name
            for (int i = 0; i < xmlMaps.Count; i++)
            {
                dynamic map = xmlMaps[i];
                if (map.Name != null && map.Name.Equals(xmlMapName, StringComparison.OrdinalIgnoreCase))
                {
                    mapIndex = i;
                    break;
                }
            }

            // If found, clear linked cells and remove the map
            if (mapIndex != -1)
            {
                dynamic map = xmlMaps[mapIndex];
                // Clear linked cells (method may be unavailable; guard with try-catch)
                try
                {
                    map.ClearAllLinkedCells();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unable to clear linked cells: {ex.Message}");
                }

                // Remove the map
                xmlMaps.RemoveAt(mapIndex);
                Console.WriteLine($"XML map \"{xmlMapName}\" removed successfully.");
            }
            else
            {
                Console.WriteLine($"XML map \"{xmlMapName}\" not found.");
            }

            // Ensure output directory exists and save workbook
            EnsureOutputDirectory(outputPath);
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper to create output directory if needed
    private static void EnsureOutputDirectory(string outputPath)
    {
        string outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }
    }
}
