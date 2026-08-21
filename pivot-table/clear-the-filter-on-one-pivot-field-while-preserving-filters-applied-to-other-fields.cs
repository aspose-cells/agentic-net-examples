// Title: Clear a single pivot field filter while preserving others with Aspose.Cells for .NET (C#)
// Description: Shows how to build a workbook, add a pivot table, apply label filters to the Category and SubCategory rows, then remove only the SubCategory filter using PivotFilters.ClearFilter, refresh the pivot, and confirm the Category filter stays active—all in C# with Aspose.Cells.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | PivotFilters.ClearFilter | clear specific pivot filter | remove one pivot field filter | preserve other pivot filters | programmatic Excel pivot filter | filter management Aspose
// Common Searches: Aspose.Cells clear filter on one pivot field only | C# remove specific pivot filter without affecting others | how to use PivotFilters.ClearFilter in Aspose.Cells | clear subcategory filter in Excel pivot table C# | preserve existing pivot filters after clearing one
// Developer Intent: Remove the filter applied to a chosen pivot field while keeping all other pivot filters unchanged.
// Use Cases: User changes a sub‑category selection and the app needs to reset that filter without losing the already‑selected category filter. | Generate a report where a temporary date filter is cleared before exporting the pivot table to Excel. | Dynamically update a dashboard by programmatically clearing only the region filter while retaining product and time filters.
// AI Prompts: Write C# code that clears the filter on pivot field index 2 using Aspose.Cells without affecting other filters. | Explain how to verify the remaining PivotFilters count after calling ClearFilter on a specific field. | Provide a C# example that iterates through PivotFilters and removes filters based on field names instead of indexes.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Shows how to build a workbook, add a pivot table, apply label filters to the Category and SubCategory rows, then remove only the SubCategory filter using PivotFilters.ClearFilter, refresh the pivot, and confirm the Category filter stays active—all in C# with Aspose.Cells.
    public class ClearSpecificPivotFieldFilterDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                // Columns: Category, SubCategory, Sales
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("SubCategory");
                worksheet.Cells["C1"].PutValue("Sales");

                worksheet.Cells["A2"].PutValue("Fruit");
                worksheet.Cells["B2"].PutValue("Apple");
                worksheet.Cells["C2"].PutValue(120);

                worksheet.Cells["A3"].PutValue("Fruit");
                worksheet.Cells["B3"].PutValue("Banana");
                worksheet.Cells["C3"].PutValue(80);

                worksheet.Cells["A4"].PutValue("Vegetable");
                worksheet.Cells["B4"].PutValue("Carrot");
                worksheet.Cells["C4"].PutValue(60);

                worksheet.Cells["A5"].PutValue("Vegetable");
                worksheet.Cells["B5"].PutValue("Broccoli");
                worksheet.Cells["C5"].PutValue(90);

                // Add a pivot table based on the data range
                int pivotIndex = worksheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

                // Add row fields: Category (field index 0) and SubCategory (field index 1)
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0); // Category
                pivotTable.AddFieldToArea(PivotFieldType.Row, 1); // SubCategory

                // Add data field: Sales (field index 2)
                pivotTable.AddFieldToArea(PivotFieldType.Data, 2);

                // Apply filters on both row fields
                // Filter Category to show only "Fruit"
                pivotTable.PivotFilters.AddLabelFilter(0, PivotFilterType.CaptionEqual, "Fruit", null);
                // Filter SubCategory to show only "Apple"
                pivotTable.PivotFilters.AddLabelFilter(1, PivotFilterType.CaptionEqual, "Apple", null);

                // Refresh and calculate to apply the filters
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Clear the filter only on SubCategory (field index 1) while keeping the Category filter.
                pivotTable.PivotFilters.ClearFilter(1);

                // Refresh again to reflect the change
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Output the remaining filter count to verify only one filter remains
                Console.WriteLine($"Remaining filters count: {pivotTable.PivotFilters.Count}");
                // Expected output: 1 (the filter on Category)

                // Save the workbook
                string outputPath = "ClearSpecificPivotFieldFilterDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
