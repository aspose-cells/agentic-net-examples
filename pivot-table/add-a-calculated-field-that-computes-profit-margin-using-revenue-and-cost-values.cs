// Title: Add a profit‑margin calculated field to an Aspose.Cells pivot table in C# and format it as a percentage
// AI Prompts: Generate C# code that creates a PivotTable with Revenue and Cost columns, inserts a calculated field named ProfitMargin using the formula (Revenue‑Cost)/Revenue, applies a 0.00% number format, refreshes the pivot, and saves the workbook. | Show how to use Aspose.Cells for .NET to add a custom calculated field to an existing pivot table, set its number format to percentage, recalculate the pivot data, and export the file.
// Common Searches: asp.net aspose.cells add profit margin calculated field to pivot table | c# Aspose.Cells pivot table calculated field percentage format example | how to refresh Aspose.Cells pivot after adding a custom calculated field
// Tags: add calculated field to Aspose.Cells pivot table | profit margin calculated field Aspose.Cells | percentage format for pivot calculated field | refresh pivot after adding calculated field Aspose.Cells | Aspose.Cells pivot table C# example

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCalculatedField
{
    // The sample creates a workbook, fills it with Product, Revenue, and Cost data, builds a pivot table, adds a calculated field called ProfitMargin using the formula (Revenue‑Cost)/Revenue, formats this field as a percentage, refreshes and calculates the pivot, and saves the result as PivotTable_ProfitMargin.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data: Product, Revenue, Cost
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Revenue";
            cells["C1"].Value = "Cost";

            string[] products = { "A", "B", "C", "A", "B", "C" };
            double[] revenues = { 200, 300, 250, 180, 320, 260 };
            double[] costs = { 120, 150, 130, 110, 140, 150 };

            for (int i = 0; i < products.Length; i++)
            {
                int row = i + 2;
                cells[$"A{row}"].Value = products[i];
                cells[$"B{row}"].Value = revenues[i];
                cells[$"C{row}"].Value = costs[i];
            }

            // Create a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C7", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Revenue");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Cost");

            // Add a calculated field for profit margin and drag it to the data area
            // Formula: (Revenue - Cost) / Revenue
            pivotTable.AddCalculatedField("ProfitMargin", "=(Revenue-Cost)/Revenue", true);

            // Format the calculated field as percentage
            PivotField profitMarginField = pivotTable.DataFields[pivotTable.DataFields.Count - 1];
            profitMarginField.NumberFormat = "0.00%";

            // Refresh and calculate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_ProfitMargin.xlsx");
        }
    }
}
