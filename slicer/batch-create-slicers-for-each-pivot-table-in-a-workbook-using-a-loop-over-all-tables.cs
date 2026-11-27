// Title: Create slicers for every pivot table in an Excel workbook with Aspose.Cells for .NET
// Common Searches: asp.net add slicer to each pivot table in a workbook using Aspose.Cells | c# loop through worksheets to generate slicers for all pivot tables with Aspose.Cells | how to position slicers with column offset when adding them via Aspose.Cells | batch create slicers for multiple pivot tables in an Excel file using Aspose.Cells .NET | programmatically connect a slicer to its pivot table in C# with Aspose.Cells
// Tags: Aspose.Cells add slicer to pivot table | C# loop create slicers for multiple pivot tables | Aspose.Cells set slicer column offset | batch slicer generation Excel Aspose.Cells | Aspose.Cells connect slicer to pivot table programmatically

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace BatchSlicerCreator
{
    // The example loads an existing Excel workbook, iterates over each worksheet, and for every pivot table adds a slicer at row 0 with a column offset of five columns per slicer. Each slicer receives a unique name, is automatically linked to its pivot table, and the modified workbook is saved to the specified output file.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths.
            // You can pass them as command‑line arguments or edit the defaults below.
            string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
            string outputPath = args.Length > 1 ? args[1] : "output_with_slicers.xlsx";

            try
            {
                // Load the workbook that already contains one or more pivot tables.
                Workbook workbook = new Workbook(inputPath);

                // Iterate through every worksheet in the workbook.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the collection of pivot tables on the current worksheet.
                    PivotTableCollection pivots = sheet.PivotTables;

                    // If there are no pivot tables, continue to the next sheet.
                    if (pivots == null || pivots.Count == 0)
                        continue;

                    // Loop over each pivot table and create a slicer for it.
                    for (int i = 0; i < pivots.Count; i++)
                    {
                        PivotTable pivot = pivots[i];

                        // Determine a location for the slicer.
                        // Here we place slicers starting from row 0 and offset each slicer
                        // by 5 columns to avoid overlap.
                        int slicerRow = 0;
                        int slicerColumn = i * 5;

                        // Give each slicer a unique name.
                        string slicerName = $"Slicer_{sheet.Name}_{i + 1}";

                        // Add the slicer to the worksheet.
                        // The overload used is: Add(PivotTable pivotTable, int row, int column, string name)
                        int slicerIndex = sheet.Slicers.Add(pivot, slicerRow, slicerColumn, slicerName);

                        // The slicer is automatically connected to the pivot table,
                        // but you can explicitly add the connection if needed:
                        // sheet.Slicers[slicerIndex].AddPivotConnection(pivot);
                    }
                }

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
