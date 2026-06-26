using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerRemoval
{
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    SlicerCollection slicers = sheet.Slicers;
                    List<Slicer> slicersToRemove = new List<Slicer>();

                    // Examine each slicer
                    foreach (Slicer slicer in slicers)
                    {
                        // Check each pivot table on the same worksheet
                        foreach (PivotTable pivot in sheet.PivotTables)
                        {
                            try
                            {
                                // Attempt to remove the connection; if not connected, an exception is thrown
                                slicer.RemovePivotConnection(pivot);

                                // Calculate row count of the pivot table using its TableRange2
                                int rowCount = pivot.TableRange2.EndRow - pivot.TableRange2.StartRow + 1;

                                // Mark slicer for removal if pivot has more than 100 rows
                                if (rowCount > 100)
                                {
                                    slicersToRemove.Add(slicer);
                                }

                                // No need to re‑establish the connection if we plan to delete the slicer
                            }
                            catch
                            {
                                // Slicer not connected to this pivot table; continue
                            }
                        }
                    }

                    // Remove identified slicers
                    foreach (Slicer s in slicersToRemove)
                    {
                        slicers.Remove(s);
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}