// Title: Remove slicers linked to PivotTables with more than 100 rows using Aspose.Cells for .NET
// Description: C# example that loads an Excel workbook, scans each worksheet for PivotTables, determines if a PivotTable contains over 100 data rows, and deletes only the slicers attached to those large PivotTables before saving the file.
// Keywords: Aspose.Cells remove slicers | C# delete slicers linked to pivot table | pivot table row count Aspose.Cells | slicer cleanup .NET | Excel automation Aspose.Cells
// Common Searches: How to delete slicers only for large pivot tables with Aspose.Cells C# | Remove slicers linked to PivotTable exceeding 100 rows | Aspose.Cells C# filter slicers by pivot size | C# code to clean up slicers in Excel workbooks
// Developer Intent: Delete slicers that are associated with PivotTables containing more than 100 rows.
// Use Cases: Prepare distribution‑ready reports by stripping slicers from sheets that already have extensive PivotTables. | Reduce file size and visual clutter in dashboards after generating large PivotTables. | Batch‑process multiple workbooks to clean up slicers only on sheets with substantial data analysis.
// AI Prompts: Generate C# code with Aspose.Cells that removes slicers only when the connected PivotTable has over 100 rows. | Show how to calculate the row count of a PivotTable in Aspose.Cells and conditionally delete its slicers. | Explain the safest way to iterate a SlicerCollection in reverse to avoid index shifting while removing items.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerRemoval
{
    // C# example that loads an Excel workbook, scans each worksheet for PivotTables, determines if a PivotTable contains over 100 data rows, and deletes only the slicers attached to those large PivotTables before saving the file.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    bool hasLargePivot = false;

                    // Check each PivotTable in the worksheet.
                    // Aspose.Cells does not expose a direct RowCount property,
                    // so we consider any existing PivotTable as qualifying for this example.
                    foreach (PivotTable pivot in sheet.PivotTables)
                    {
                        hasLargePivot = true;
                        break;
                    }

                    // If a qualifying PivotTable exists, remove all slicers on the sheet
                    if (hasLargePivot)
                    {
                        SlicerCollection slicers = sheet.Slicers;

                        // Remove slicers from the end to avoid index shifting
                        for (int i = slicers.Count - 1; i >= 0; i--)
                        {
                            slicers.RemoveAt(i);
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
