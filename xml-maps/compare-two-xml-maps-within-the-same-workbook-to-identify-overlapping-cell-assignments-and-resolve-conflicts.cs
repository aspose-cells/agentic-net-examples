// Title: Identify and resolve overlapping cell assignments between two XML maps in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, extracts the cell addresses from XmlMap1 and XmlMap2, lists the addresses that appear in both maps, and clears those cells in the second map. | Generate a method that accepts a Workbook object and two XML map names, returns the conflicting cell addresses, and removes the values from those cells for the specified map.
// Common Searches: how to detect duplicate cell references in two XML maps with Aspose.Cells C# | Aspose.Cells example for comparing XmlMap1 and XmlMap2 cell assignments | C# code to clear overlapping cells of an XML map in an Excel file | resolve XML map cell conflicts in a .xlsx workbook using Aspose.Cells
// Tags: xmlmap overlapping cell detection Aspose.Cells | clear conflicting xmlmap cells C# | compare xmlmap cell addresses workbook | aspnet excel xmlmap conflict resolution | importedxmlmapitems cell enumeration Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace XmlMapConflictResolver
{
    // The sample loads an Excel workbook, retrieves two XML maps (XmlMap1 and XmlMap2), gathers all cell addresses referenced by each map via ImportedXmlMapItems, finds the addresses that are assigned to both maps, prints the overlapping cells, clears the values of those cells for the second map across all worksheets, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Input and output file paths
                string inputPath = "InputWorkbook.xlsx";
                string outputPath = "OutputWorkbook.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from file
                Workbook workbook = new Workbook(inputPath);

                // Use dynamic to access XmlMaps (avoids compile‑time binding issues with different library versions)
                dynamic xmlMaps = workbook;
                dynamic xmlMap1 = null;
                dynamic xmlMap2 = null;

                try
                {
                    xmlMap1 = xmlMaps.XmlMaps["XmlMap1"];
                    xmlMap2 = xmlMaps.XmlMaps["XmlMap2"];
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error accessing XML maps: {ex.Message}");
                    return;
                }

                if (xmlMap1 == null || xmlMap2 == null)
                {
                    Console.WriteLine("One or both XML maps were not found.");
                    return;
                }

                // Helper to collect all cell addresses used by an XML map
                HashSet<string> GetMappedCellAddresses(dynamic map)
                {
                    var addresses = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    try
                    {
                        foreach (var item in map.ImportedXmlMapItems)
                        {
                            // ImportedXmlMapItem may expose CellName property
                            string cellName = item.CellName as string;
                            if (!string.IsNullOrEmpty(cellName))
                            {
                                addresses.Add(cellName);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to read ImportedXmlMapItems: {ex.Message}");
                    }
                    return addresses;
                }

                // Get cell address sets for both maps
                HashSet<string> map1Cells = GetMappedCellAddresses(xmlMap1);
                HashSet<string> map2Cells = GetMappedCellAddresses(xmlMap2);

                // Identify overlapping cells
                var overlappingCells = new List<string>();
                foreach (var cell in map1Cells)
                {
                    if (map2Cells.Contains(cell))
                    {
                        overlappingCells.Add(cell);
                    }
                }

                // Output overlapping cells
                Console.WriteLine("Overlapping cell assignments:");
                foreach (var cell in overlappingCells)
                {
                    Console.WriteLine(cell);
                }

                // Resolve conflicts: clear values of overlapping cells for the second map
                foreach (var cellAddress in overlappingCells)
                {
                    foreach (Worksheet ws in workbook.Worksheets)
                    {
                        try
                        {
                            Cell cell = ws.Cells[cellAddress];
                            if (cell != null)
                            {
                                cell.PutValue(string.Empty);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error clearing cell {cellAddress} in sheet {ws.Name}: {ex.Message}");
                        }
                    }
                }

                // Save the workbook with resolved conflicts
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
