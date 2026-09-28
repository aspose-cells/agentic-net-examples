// Title: List all XML map names in an Excel workbook and print them to the console using Aspose.Cells for .NET
// AI Prompts: Generate a C# console program that loads an .xlsx file with Aspose.Cells, uses reflection to obtain the XmlMaps collection if it exists, and writes each XmlMap's Name to the console. | Create a .NET snippet that checks whether the current Aspose.Cells version supports XML maps, iterates over all XmlMap objects safely, and outputs their names while handling missing or empty collections. | Write C# code that gracefully falls back when the XmlMaps property is unavailable, enumerates any existing XML maps in a workbook, and prints each map's identifier to standard output.
// Common Searches: c# retrieve xml map names from workbook using aspose.cells | how to enumerate XmlMaps collection with reflection in Aspose.Cells | list xml maps in an xlsx file when XmlMaps property is missing | asp.net console output of Excel XML map identifiers | fallback method for accessing XmlMaps in older Aspose.Cells versions
// Tags: Aspose.Cells enumerate XmlMap objects | C# retrieve XmlMaps collection via reflection | console output of Excel XML map names | fallback handling for unsupported XmlMaps property | handle empty or missing XML maps Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // This example loads an Excel workbook with Aspose.Cells, uses reflection to access the XmlMaps collection for version‑agnostic compatibility, iterates through each XmlMap, extracts its Name property, and prints the names to the console. It also includes graceful handling for cases where no XML maps are present or the XmlMaps feature is unavailable in the current library version.
    class Program
    {
        static void Main()
        {
            string filePath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Attempt to retrieve the XmlMaps collection via reflection (covers versions where XmlMaps may be unavailable)
                PropertyInfo xmlMapsProp = workbook.GetType().GetProperty("XmlMaps", BindingFlags.Public | BindingFlags.Instance);
                if (xmlMapsProp != null)
                {
                    object xmlMapsObj = xmlMapsProp.GetValue(workbook);
                    if (xmlMapsObj is System.Collections.IEnumerable xmlMapsEnumerable)
                    {
                        bool anyMap = false;
                        foreach (object mapObj in xmlMapsEnumerable)
                        {
                            anyMap = true;
                            // Retrieve the Name property of each XmlMap
                            PropertyInfo nameProp = mapObj.GetType().GetProperty("Name", BindingFlags.Public | BindingFlags.Instance);
                            string mapName = nameProp?.GetValue(mapObj)?.ToString() ?? "<Unnamed>";
                            Console.WriteLine(mapName);
                        }

                        if (!anyMap)
                        {
                            Console.WriteLine("No XML maps found in the workbook.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("XML maps collection is empty or not enumerable.");
                    }
                }
                else
                {
                    Console.WriteLine("The current Aspose.Cells version does not support XML maps.");
                }
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
