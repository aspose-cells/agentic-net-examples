// Title: Create a Pivot Table on a New Sheet with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to build a workbook, fill a source sheet with Category, Product, and Sales data, use MaxDisplayRange to define the source range, add a new worksheet, insert a pivot table at A1, assign row, column, and data fields, apply PivotTableStyleMedium9, refresh the pivot, and save the file.
// Keywords: Aspose.Cells | C# pivot table | Aspose.Cells PivotTable | create pivot table programmatically | add pivot table to new worksheet | MaxDisplayRange | PivotTableStyleMedium9 | .NET Excel automation | Aspose.Cells example
// Common Searches: Aspose.Cells create pivot table on another sheet | C# generate pivot table from range using Aspose.Cells | How to set pivot table style with Aspose.Cells | Refresh pivot tables Aspose.Cells .NET | Add row column data fields Aspose.Cells pivot
// Developer Intent: Generate a pivot table from source data on one worksheet and place it on a newly added worksheet using Aspose.Cells for .NET.
// Use Cases: Summarize sales data by Category (rows) and Product (columns) in a separate worksheet. | Apply a built‑in pivot table style for a polished report layout. | Refresh pivot tables programmatically to ensure calculations are up‑to‑date before saving. | Reuse the same source range to create multiple pivot tables across different sheets.
// AI Prompts: Write C# code with Aspose.Cells that creates a pivot table from a worksheet range, adds it to a new sheet, configures row, column, and data fields, applies a style, refreshes, and saves the workbook. | Explain how MaxDisplayRange is used to build the source data reference string for a pivot cache in Aspose.Cells. | Provide troubleshooting steps for handling exceptions when generating pivot tables with Aspose.Cells and logging errors.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsPivotExample
{
    // Demonstrates how to build a workbook, fill a source sheet with Category, Product, and Sales data, use MaxDisplayRange to define the source range, add a new worksheet, insert a pivot table at A1, assign row, column, and data fields, apply PivotTableStyleMedium9, refresh the pivot, and save the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // 1. Prepare source data on the first worksheet
                // -------------------------------------------------
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "SourceData";

                // Populate header row
                sourceSheet.Cells["A1"].PutValue("Category");
                sourceSheet.Cells["B1"].PutValue("Product");
                sourceSheet.Cells["C1"].PutValue("Sales");

                // Sample data
                string[] categories = { "Fruit", "Fruit", "Vegetable", "Vegetable", "Fruit" };
                string[] products   = { "Apple", "Banana", "Carrot", "Potato", "Orange" };
                int[] sales         = { 120, 80, 150, 200, 90 };

                for (int i = 0; i < categories.Length; i++)
                {
                    int row = i + 2; // data starts at row 2
                    sourceSheet.Cells[$"A{row}"].PutValue(categories[i]);
                    sourceSheet.Cells[$"B{row}"].PutValue(products[i]);
                    sourceSheet.Cells[$"C{row}"].PutValue(sales[i]);
                }

                // -------------------------------------------------
                // 2. Create a new worksheet that will host the pivot table
                // -------------------------------------------------
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // -------------------------------------------------
                // 3. Define the source data range for the pivot cache
                // -------------------------------------------------
                AsposeRange usedRange = sourceSheet.Cells.MaxDisplayRange;
                // Build the external reference string, e.g. =SourceData!A1:C6
                string sourceData = $"=SourceData!{usedRange.Address}";

                // -------------------------------------------------
                // 4. Add the pivot table to the new worksheet
                // -------------------------------------------------
                PivotTableCollection pivotTables = pivotSheet.PivotTables;
                int pivotIndex = pivotTables.Add(sourceData, "A1", "SalesPivot");
                PivotTable pivotTable = pivotTables[pivotIndex];

                // -------------------------------------------------
                // 5. Configure the pivot table fields
                // -------------------------------------------------
                // Row field: Category
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                // Column field: Product
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
                // Data field: Sum of Sales
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Optional: set a style for better appearance
                pivotTable.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

                // -------------------------------------------------
                // 6. Refresh the pivot table to calculate data
                // -------------------------------------------------
                pivotSheet.RefreshPivotTables();

                // -------------------------------------------------
                // 7. Save the workbook
                // -------------------------------------------------
                string outputPath = "PivotTableOnNewSheet.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
