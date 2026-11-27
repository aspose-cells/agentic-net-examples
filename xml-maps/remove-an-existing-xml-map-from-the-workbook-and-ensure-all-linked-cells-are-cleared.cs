// Title: Remove an XML map from an Excel workbook and clear all linked cells with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# remove specific XML map from workbook | Delete XML map and clear bound cells in Excel using Aspose.Cells | How to clear cells linked to a removed XML map in Aspose.Cells | C# code to delete XML map by name and save workbook without it | Remove XML map and reset data bindings with Aspose.Cells .NET
// Tags: remove xml map Aspose.Cells | clear linked cells Aspose.Cells | xml map deletion C# | excel workbook without xml map Aspose | reset data bindings Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace RemoveXmlMapDemo
{
    // The example loads an Excel workbook, finds the XML map by its name, removes the map from the XmlMaps collection, clears the contents of all used cells in each worksheet to eliminate data bindings, and saves the modified workbook as a new file.
    class Program
    {
        /// <param name="args">
        /// args[0] - Path to the source workbook (e.g., "input.xlsx").
        /// args[1] - Name of the XML map to remove (e.g., "MyMap").
        /// </param>
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: RemoveXmlMapDemo <workbookPath> <xmlMapName>");
                return;
            }

            string workbookPath = args[0];
            string xmlMapName   = args[1];

            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);
                XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

                // Locate the XML map by name
                int mapIndex = -1;
                for (int i = 0; i < xmlMaps.Count; i++)
                {
                    if (string.Equals(xmlMaps[i].Name, xmlMapName, StringComparison.OrdinalIgnoreCase))
                    {
                        mapIndex = i;
                        break;
                    }
                }

                if (mapIndex == -1)
                {
                    Console.WriteLine($"XML map '{xmlMapName}' not found in the workbook.");
                }
                else
                {
                    // Remove the XML map from the collection
                    xmlMaps.RemoveAt(mapIndex);
                    Console.WriteLine($"XML map '{xmlMapName}' removed successfully.");

                    // Clear all cells that were linked to the removed map.
                    // Clear the used range of each worksheet to avoid wiping empty rows/columns.
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        int maxRow = sheet.Cells.MaxDataRow;
                        int maxCol = sheet.Cells.MaxDataColumn;

                        // If the sheet is empty, MaxDataRow/MaxDataColumn return -1.
                        if (maxRow >= 0 && maxCol >= 0)
                        {
                            // Clear values, formulas, and data bindings in the used range.
                            sheet.Cells.ClearContents(0, 0, maxRow, maxCol);
                        }
                    }

                    Console.WriteLine("All linked cells have been cleared.");
                }

                // Determine output path
                string outputDir = Path.GetDirectoryName(workbookPath) ?? Environment.CurrentDirectory;
                string outputPath = Path.Combine(outputDir,
                    Path.GetFileNameWithoutExtension(workbookPath) + "_NoXmlMap.xlsx");

                // Ensure the output directory exists
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Modified workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
