// Title: Remove slicers linked to pivot tables with more than 100 rows using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells remove slicer when linked pivot table exceeds 100 rows | Delete Excel slicers based on pivot table size with Aspose.Cells .NET | How to filter slicers by pivot table row count using Aspose.Cells in C# | Programmatically remove slicers attached to large pivot tables with Aspose.Cells | Iterate worksheets and delete slicers linked to big pivot tables in C#
// Tags: Aspose.Cells remove slicer by pivot row count | C# slicer collection manipulation Aspose.Cells | pivot table DataBodyRange evaluation Aspose.Cells | Excel workbook slicer deletion .NET | filter slicers linked to large pivot tables C# | Aspose.Cells worksheet iteration for slicer removal

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace RemoveLargePivotSlicers
{
    // The example loads an Excel workbook, walks through each worksheet, checks every slicer against the pivot tables on that sheet, and removes the slicer if any linked pivot table's DataBodyRange contains more than 100 rows, then saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths (use defaults if not provided)
            string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
            string outputPath = args.Length > 1 ? args[1] : "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook \"{inputPath}\": {ex.Message}");
                return;
            }

            // Iterate through all worksheets
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                SlicerCollection slicers = worksheet.Slicers;
                if (slicers == null || slicers.Count == 0)
                    continue;

                // Get all pivot tables on the worksheet once
                PivotTableCollection worksheetPivots = worksheet.PivotTables;

                // Iterate backwards so that removal does not affect the index order
                for (int s = slicers.Count - 1; s >= 0; s--)
                {
                    Slicer slicer = slicers[s];
                    bool shouldRemove = false;

                    // Evaluate each pivot table on the worksheet
                    foreach (PivotTable pivot in worksheetPivots)
                    {
                        try
                        {
                            // DataBodyRange may be null for empty pivots; guard against it
                            CellArea? dataArea = pivot.DataBodyRange;
                            if (dataArea.HasValue)
                            {
                                int rowCount = dataArea.Value.EndRow - dataArea.Value.StartRow + 1;
                                if (rowCount > 100)
                                {
                                    shouldRemove = true;
                                    break;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error evaluating pivot \"{pivot.Name}\": {ex.Message}");
                            // Continue with next pivot
                        }
                    }

                    // Remove the slicer if any linked pivot table exceeds 100 rows
                    if (shouldRemove)
                    {
                        slicers.RemoveAt(s);
                        Console.WriteLine($"Removed slicer \"{slicer.Name}\" from worksheet \"{worksheet.Name}\".");
                    }
                }
            }

            // Ensure the output directory exists
            try
            {
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to prepare output directory: {ex.Message}");
                return;
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
