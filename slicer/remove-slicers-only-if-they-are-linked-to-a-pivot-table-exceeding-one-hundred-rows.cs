// Title: Delete slicers linked to pivot tables that have more than 100 data rows using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that iterates over every slicer on a worksheet, determines the row count of the pivot table each slicer is attached to, and deletes the slicer when the pivot source contains over 100 data rows. | Demonstrate how to associate a slicer with its pivot table, calculate the number of data rows from the source range, and perform conditional slicer removal before saving the workbook.
// Common Searches: aspocells c# remove slicer if pivot table has more than 100 rows | how to delete Excel slicers based on pivot source size using Aspose.Cells | conditional slicer cleanup in .NET workbook with large pivot tables | C# example for filtering slicers by pivot data row count Aspose.Cells | programmatically remove slicers linked to big pivot tables in Excel via Aspose
// Tags: Aspose.Cells slicer deletion based on pivot rows | C# filter slicers by pivot source size | Excel slicer cleanup with Aspose.Cells API | pivot table row count threshold for slicer handling | programmatic slicer management in .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerRemovalDemo
{
    // The sample creates a workbook with 101 data rows, adds a pivot table and a linked slicer, maps the slicer to its pivot source range, checks if the source contains more than 100 data rows, removes the slicer when the condition is met, and saves the resulting Excel file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (101 rows to exceed the 100‑row threshold)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                for (int i = 2; i <= 102; i++) // rows 2..102 => 101 data rows
                {
                    sheet.Cells[$"A{i}"].PutValue("Item" + (i - 1));
                    sheet.Cells[$"B{i}"].PutValue(i * 10);
                }

                // Define the source range for the pivot table
                string sourceRange = "A1:B102";

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add(sourceRange, "D1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIndex];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Value");

                // Refresh pivot cache and calculate data
                pivot.RefreshData();          // corrected API usage
                pivot.CalculateData();

                // Add a slicer linked to the pivot table
                int slicerIndex = sheet.Slicers.Add(pivot, "F1", "Category");
                Slicer slicer = sheet.Slicers[slicerIndex];

                // Map slicer to its pivot table and store the source range
                var slicerPivotMap = new Dictionary<Slicer, (PivotTable Pivot, string SourceRange)>
                {
                    { slicer, (pivot, sourceRange) }
                };

                // Identify slicers whose associated pivot source has more than 100 data rows
                var slicersToRemove = new List<Slicer>();
                foreach (Slicer s in sheet.Slicers)
                {
                    if (slicerPivotMap.TryGetValue(s, out var info))
                    {
                        string rangePart = info.SourceRange; // e.g., "A1:B102"

                        // Split the range into start and end cells
                        string[] parts = rangePart.Split(':');
                        if (parts.Length == 2)
                        {
                            // Create CellArea from start and end cells
                            CellArea area = CellArea.CreateCellArea(parts[0], parts[1]);
                            int sourceRows = area.EndRow - area.StartRow + 1; // includes header
                            int dataRows = sourceRows - 1; // exclude header

                            if (dataRows > 100)
                            {
                                slicersToRemove.Add(s);
                            }
                        }
                    }
                }

                // Remove identified slicers
                foreach (Slicer s in slicersToRemove)
                {
                    sheet.Slicers.Remove(s);
                }

                // Save the workbook
                string outputPath = "SlicerRemovalResult.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
