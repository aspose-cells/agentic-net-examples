// Title: Create a PivotTable on a New Sheet from Existing Data with Aspose.Cells for .NET
// Description: C# example that builds a workbook, fills a source sheet with Category, Product and Sales data, defines the used range, adds a second worksheet, creates a PivotTable named SalesPivot at cell A3, assigns Category to rows, Product to columns, Sales to data (sum), applies a medium style, refreshes the pivot, and saves the file as PivotTableFromExistingTable.xlsx.
// Keywords: Aspose.Cells pivot table C# | add pivot table new worksheet .NET | generate pivot from existing range Aspose | PivotTableStyleType Aspose.Cells | refresh pivot tables programmatically | Excel pivot table automation | Aspose.Cells example pivot
// Common Searches: how to create a pivot table on a different sheet using Aspose.Cells | Aspose.Cells C# pivot table from source range | set row, column, data fields in Aspose.Cells pivot | apply style to Aspose.Cells PivotTable | refresh all pivot tables Aspose.Cells
// Developer Intent: Generate a PivotTable from data on one worksheet and place it on a separate worksheet using Aspose.Cells for .NET.
// Use Cases: Produce a sales summary report that groups totals by Category (rows) and Product (columns) on a dedicated reporting sheet. | Separate raw data and analytical views by creating pivot tables on new worksheets automatically. | Apply a predefined PivotTableStyleType for better readability and refresh the pivot after data updates.
// AI Prompts: Write C# code with Aspose.Cells to create a PivotTable from a used range and place it on a new worksheet. | Explain how to add row, column, and data fields to a PivotTable and set its style using Aspose.Cells for .NET. | Show how to refresh all PivotTables on a worksheet after modifying source data with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsPivotExample
{
    // C# example that builds a workbook, fills a source sheet with Category, Product and Sales data, defines the used range, adds a second worksheet, creates a PivotTable named SalesPivot at cell A3, assigns Category to rows, Product to columns, Sales to data (sum), applies a medium style, refreshes the pivot, and saves the file as PivotTableFromExistingTable.xlsx.
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
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "SourceData";

                // Populate sample data (Category, Product, Sales)
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Product");
                dataSheet.Cells["C1"].PutValue("Sales");

                string[] categories = { "Fruit", "Fruit", "Vegetable", "Vegetable", "Fruit" };
                string[] products   = { "Apple", "Banana", "Carrot", "Potato", "Orange" };
                int[] sales         = { 1200, 800, 600, 900, 1500 };

                for (int i = 0; i < categories.Length; i++)
                {
                    dataSheet.Cells[i + 2, 0].PutValue(categories[i]); // Column A
                    dataSheet.Cells[i + 2, 1].PutValue(products[i]);   // Column B
                    dataSheet.Cells[i + 2, 2].PutValue(sales[i]);     // Column C
                }

                // -------------------------------------------------
                // 2. Add a new worksheet that will host the pivot table
                // -------------------------------------------------
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // -------------------------------------------------
                // 3. Define the source data range for the pivot cache
                // -------------------------------------------------
                // Use MaxDisplayRange to get the used range automatically
                AsposeRange sourceRange = dataSheet.Cells.MaxDisplayRange;
                string sourceData = $"=SourceData!{sourceRange.Address}";

                // -------------------------------------------------
                // 4. Add the pivot table to the new worksheet
                // -------------------------------------------------
                // Destination cell for the pivot table (top‑left corner)
                string destCell = "A3";
                string pivotName = "SalesPivot";

                // Add the pivot table
                int pivotIndex = pivotSheet.PivotTables.Add(sourceData, destCell, pivotName);
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

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
                string outputPath = "PivotTableFromExistingTable.xlsx";
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
