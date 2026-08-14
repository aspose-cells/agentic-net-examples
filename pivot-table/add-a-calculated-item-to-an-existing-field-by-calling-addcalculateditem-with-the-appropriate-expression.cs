// Title: C# – Add a Calculated Item to a Pivot Table Row Field with Aspose.Cells
// Description: Learn how to create an Excel workbook in C#, build a simple data set, insert a pivot table, and use Aspose.Cells' PivotField.AddCalculatedItem method to add a custom row item (e.g., "Food_Drink_Total" = Food + Drink). The example refreshes and saves the pivot as an XLSX file.
// Keywords: Aspose.Cells AddCalculatedItem C# | pivot table calculated item .NET | Aspose.Cells custom pivot row | C# Excel pivot table example | Aspose.Cells calculated field | Add calculated item to pivot table | Aspose.Cells tutorial | Excel pivot custom total
// Common Searches: Aspose.Cells add calculated item pivot table C# | how to create a calculated row item in Aspose.Cells | C# example for PivotField.AddCalculatedItem | custom total in Aspose.Cells pivot table | Aspose.Cells calculated item expression syntax
// Developer Intent: Insert a new calculated item into an existing pivot‑table row field by calling AddCalculatedItem with a valid expression.
// Use Cases: Combine several category entries into a single summarized row for reporting. | Define a reusable custom metric (e.g., sum of specific items) without modifying source data. | Provide dynamic totals that update automatically when the underlying pivot data changes.
// AI Prompts: Generate C# code that adds a calculated item using an IF statement to a pivot table with Aspose.Cells. | Show how to retrieve and format the value of a calculated item after refreshing the pivot table. | Explain how to use AddCalculatedItem to compute an average of multiple fields in a pivot table.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsCalculatedItemDemo
{
    // Learn how to create an Excel workbook in C#, build a simple data set, insert a pivot table, and use Aspose.Cells' PivotField.AddCalculatedItem method to add a custom row item (e.g., "Food_Drink_Total" = Food + Drink). The example refreshes and saves the pivot as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["A2"].Value = "Food";
            sheet.Cells["A3"].Value = "Drink";
            sheet.Cells["A4"].Value = "Food";
            sheet.Cells["A5"].Value = "Drink";

            sheet.Cells["B1"].Value = "Amount";
            sheet.Cells["B2"].Value = 120;
            sheet.Cells["B3"].Value = 80;
            sheet.Cells["B4"].Value = 150;
            sheet.Cells["B5"].Value = 70;

            // Add a pivot table covering the data range and place it at D1
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the "Category" field to the row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

            // Add the "Amount" field to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Retrieve the row field (Category) and add a calculated item
            // The calculated item will sum the values of Food and Drink
            PivotField categoryField = pivotTable.RowFields[0];
            categoryField.AddCalculatedItem("Food_Drink_Total", "=Food + Drink");

            // Refresh and calculate the pivot table to apply the new item
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook with the calculated item added
            workbook.Save("PivotTable_With_CalculatedItem.xlsx");
        }
    }
}
