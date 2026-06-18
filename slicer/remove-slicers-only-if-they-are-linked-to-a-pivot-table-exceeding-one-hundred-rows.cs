using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerRemovalDemo
{
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            try
            {
                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    SlicerCollection slicers = sheet.Slicers;

                    // Iterate backwards to safely remove items
                    for (int i = slicers.Count - 1; i >= 0; i--)
                    {
                        Slicer slicer = slicers[i];
                        bool shouldRemove = false;

                        // Check each pivot table on the same worksheet
                        foreach (PivotTable pivot in sheet.PivotTables)
                        {
                            try
                            {
                                // Attempt to remove the connection; if successful, mark for removal
                                slicer.RemovePivotConnection(pivot);
                                shouldRemove = true;
                                break;
                            }
                            catch
                            {
                                // Not connected to this pivot; continue checking others
                            }
                        }

                        // Remove slicer if it was connected to any pivot table
                        if (shouldRemove)
                        {
                            slicers.RemoveAt(i);
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during processing: {ex.Message}");
            }
        }
    }
}