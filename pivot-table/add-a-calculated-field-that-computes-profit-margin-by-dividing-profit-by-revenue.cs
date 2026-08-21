// Title: Aspose.Cells C# – Add a Profit‑Margin Calculated Field to a Pivot Table
// Description: This example shows how to create a workbook, populate it with product, revenue and profit data, build a pivot table, and insert a calculated field named **ProfitMargin** that divides Profit by Revenue. The field is formatted as a percentage, the pivot is refreshed, and the file is saved as an Excel workbook.
// Keywords: Aspose.Cells calculated field C# | pivot table profit margin | add calculated field Aspose.Cells | percentage format pivot field | refresh pivot table Aspose | C# Excel pivot example | financial analysis workbook
// Common Searches: how to add a calculated field in Aspose.Cells pivot table | profit margin formula Aspose.Cells C# | set percentage format for calculated field Aspose | refresh pivot table after adding calculated field | Aspose.Cells pivot table tutorial
// Developer Intent: Generate a pivot table and define a calculated column that computes profit margin (Profit ÷ Revenue) with percentage formatting using Aspose.Cells in C#.
// Use Cases: Create a sales‑performance report that automatically shows profit‑margin percentages per product. | Build a financial dashboard where margin values update when underlying revenue or profit figures change. | Export a ready‑to‑share Excel workbook containing a formatted profit‑margin column for stakeholder review.
// AI Prompts: Write C# code with Aspose.Cells to add a calculated field called ProfitMargin that computes Profit/Revenue and formats it as a percentage in a pivot table. | Explain the steps to refresh and recalculate an Aspose.Cells pivot table after inserting a new calculated field. | Show how to modify an existing Aspose.Cells pivot table to include a profit‑margin column and apply a 0.00% number format.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCalculatedFieldDemo
{
    // This example shows how to create a workbook, populate it with product, revenue and profit data, build a pivot table, and insert a calculated field named **ProfitMargin** that divides Profit by Revenue. The field is formatted as a percentage, the pivot is refreshed, and the file is saved as an Excel workbook.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data: Product, Revenue, Profit
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Revenue";
            cells["C1"].Value = "Profit";

            cells["A2"].Value = "A";
            cells["B2"].Value = 1000;
            cells["C2"].Value = 200;

            cells["A3"].Value = "B";
            cells["B3"].Value = 1500;
            cells["C3"].Value = 300;

            cells["A4"].Value = "C";
            cells["B4"].Value = 2000;
            cells["C4"].Value = 500;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");          // Row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Revenue");        // Data field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Profit");         // Data field

            // Add a calculated field that computes profit margin = Profit / Revenue
            // The formula must reference the source field names exactly as they appear in the data source
            pivotTable.AddCalculatedField("ProfitMargin", "=Profit/Revenue", true);

            // Optionally format the calculated field as percentage
            PivotField profitMarginField = pivotTable.DataFields[pivotTable.DataFields.Count - 1];
            profitMarginField.NumberFormat = "0.00%";

            // Refresh and calculate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_With_ProfitMargin.xlsx");
        }
    }
}
