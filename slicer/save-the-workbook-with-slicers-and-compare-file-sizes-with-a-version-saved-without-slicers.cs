// Title: C# – Compare Excel workbook size with and without slicers using Aspose.Cells
// Description: A complete C# example that creates a workbook, adds sample data, builds a pivot table, inserts a slicer, saves the file, removes the slicer, saves a second file, and prints the byte size of each XLSX to show the storage impact of slicers.
// Keywords: Aspose.Cells slicer size | C# Excel slicer example | compare XLSX file size | remove slicer Aspose.Cells | pivot table slicer impact | Aspose.Cells file size optimization | Excel slicer storage overhead
// Common Searches: Aspose.Cells C# how to measure file size with slicer | compare Excel workbook size with and without slicer | remove slicer before saving Aspose.Cells workbook | size difference XLSX when adding slicer | C# code to get file size of generated Excel file
// Developer Intent: Find out how many extra bytes a slicer adds to an XLSX file by saving the same workbook once with the slicer and once without it.
// Use Cases: Determine whether a slicer fits within attachment size limits before sending a report. | Automate workbook size optimization by stripping slicers when the file exceeds a threshold. | Validate the storage cost of slicers in generated Excel dashboards.
// AI Prompts: Generate C# code that creates a pivot table, adds a slicer, saves the workbook, removes the slicer, saves again, and outputs the size difference. | Explain how Aspose.Cells stores slicer definitions in an XLSX package and why this affects file size. | Provide a .NET method to compare two workbook files' sizes and log the result in a console application.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerSizeComparison
{
    // A complete C# example that creates a workbook, adds sample data, builds a pivot table, inserts a slicer, saves the file, removes the slicer, saves a second file, and prints the byte size of each XLSX to show the storage impact of slicers.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data
                cells["A1"].Value = "Fruit";
                cells["B1"].Value = "Year";
                cells["C1"].Value = "Amount";

                string[] fruits = { "Apple", "Banana", "Apple", "Banana", "Apple", "Banana" };
                int[] years = { 2020, 2020, 2021, 2021, 2022, 2022 };
                int[] amounts = { 100, 150, 200, 250, 300, 350 };

                for (int i = 0; i < fruits.Length; i++)
                {
                    cells[i + 1, 0].Value = fruits[i];
                    cells[i + 1, 1].Value = years[i];
                    cells[i + 1, 2].Value = amounts[i];
                }

                // Add a pivot table based on the data range
                PivotTableCollection pivots = sheet.PivotTables;
                int pivotIndex = pivots.Add("=Sheet1!A1:C7", "E3", "FruitPivot");
                PivotTable pivot = pivots[pivotIndex];
                pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
                pivot.AddFieldToArea(PivotFieldType.Column, "Year");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a slicer linked to the pivot table (filter by Fruit)
                // Note: In older Aspose.Cells versions the parameter order is (pivot, destCellName, baseFieldName)
                SlicerCollection slicers = sheet.Slicers;
                int slicerIndex = slicers.Add(pivot, "E12", "Fruit"); // destination cell, then field name
                Slicer slicer = slicers[slicerIndex];
                slicer.Caption = "Fruit Slicer";

                // Save workbook with slicer
                string withSlicerPath = "WithSlicer.xlsx";
                workbook.Save(withSlicerPath, SaveFormat.Xlsx);

                // Remove slicer(s) from the worksheet
                slicers.Clear();

                // Save workbook without slicer
                string withoutSlicerPath = "WithoutSlicer.xlsx";
                workbook.Save(withoutSlicerPath, SaveFormat.Xlsx);

                // Get file sizes
                long sizeWithSlicer = new FileInfo(withSlicerPath).Length;
                long sizeWithoutSlicer = new FileInfo(withoutSlicerPath).Length;

                // Output comparison
                Console.WriteLine($"File size with slicer    : {sizeWithSlicer} bytes");
                Console.WriteLine($"File size without slicer : {sizeWithoutSlicer} bytes");
                Console.WriteLine($"Size difference          : {sizeWithSlicer - sizeWithoutSlicer} bytes");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
