// Title: C# – Persist PivotTable Values with Aspose.Cells PivotTable.CalculateData Before Saving
// Description: Demonstrates how to create a workbook, populate sample data, add a pivot table, assign row and data fields, refresh the cache, call PivotTable.CalculateData to materialize results, and save the file so the calculated pivot values are stored permanently.
// Keywords: Aspose.Cells PivotTable.CalculateData | C# pivot table save calculated values | materialize pivot data Aspose.Cells | RefreshData vs CalculateData Aspose | pre‑calculated pivot workbook .NET | Aspose.Cells example pivot table | save workbook with pivot results
// Common Searches: Aspose.Cells calculate pivot data before saving | PivotTable.CalculateData C# example | How to materialize pivot values in Aspose.Cells | RefreshData and CalculateData difference Aspose | Save Excel file with pre‑calculated pivot using Aspose
// Developer Intent: Ensure pivot table results are computed and embedded in the workbook by invoking CalculateData prior to saving.
// Use Cases: Generate static sales reports where pivot totals are fixed at export time. | Create Excel files that open with correct pivot values without requiring user refresh. | Automate dashboard data feeds that rely on pre‑calculated pivot summaries.
// AI Prompts: Show me how to add multiple data fields to a pivot and calculate them with Aspose.Cells in C#. | Provide a C# snippet that creates a pivot, refreshes the cache, calls CalculateData, and exports the workbook to PDF. | Explain the functional differences between RefreshData and CalculateData for Aspose.Cells pivot tables.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableCalculateDataDemo
{
    // Demonstrates how to create a workbook, populate sample data, add a pivot table, assign row and data fields, refresh the cache, call PivotTable.CalculateData to materialize results, and save the file so the calculated pivot values are stored permanently.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Fruit");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["A3"].PutValue("Vegetable");
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["A4"].PutValue("Fruit");
            sheet.Cells["B4"].PutValue(70);
            sheet.Cells["A5"].PutValue("Vegetable");
            sheet.Cells["B5"].PutValue(40);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: rows = Category, data = Amount
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Refresh the pivot cache (optional but ensures up‑to‑date source)
            pivotTable.RefreshData();

            // Calculate the pivot data so that cell values are materialized
            pivotTable.CalculateData();

            // Save the workbook – the calculated pivot data will be stored in the file
            workbook.Save("PivotTableWithCalculatedData.xlsx");
        }
    }
}
