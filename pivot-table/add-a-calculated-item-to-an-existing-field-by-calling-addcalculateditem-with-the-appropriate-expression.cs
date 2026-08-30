// Title: Add a calculated item to a pivot table row field with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, builds a pivot table from a data range, and uses PivotField.AddCalculatedItem to define a new row item that sums two categories. | Show how to refresh and recalculate a pivot table after adding a calculated item using the Aspose.Cells API. | Generate a complete example that saves the workbook with the custom calculated pivot item to an .xlsx file.
// Common Searches: aspnet add calculated item to pivot table using Aspose.Cells | C# Aspose.Cells PivotField AddCalculatedItem example code | how to sum two row items in an Aspose.Cells pivot table | create custom calculated row item in Excel workbook with Aspose.Cells .NET | Aspose.Cells calculated item expression syntax for pivot tables
// Tags: Aspose.Cells calculated pivot item API C# | pivot table custom expression in Aspose.Cells | C# generate .xlsx with calculated row item | refresh and calculate pivot after adding item Aspose.Cells | define calculated item for row field Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsCalculatedItemDemo
{
    // The example creates a new workbook, fills it with sample Category and Amount data, adds a pivot table on range A1:B5, places the Category field in the row area and Amount in the data area, then adds a calculated row item named 'Food_Drink_Total' with the expression '=Food + Drink' using PivotField.AddCalculatedItem, refreshes and calculates the pivot, and finally saves the file as PivotTableWithCalculatedItem.xlsx.
    public class Program
    {
        public static void Main()
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
            // The calculated item will sum the values of "Food" and "Drink"
            PivotField categoryField = pivotTable.RowFields[0];
            categoryField.AddCalculatedItem("Food_Drink_Total", "=Food + Drink");

            // Refresh and calculate the pivot table to apply the new item
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotTableWithCalculatedItem.xlsx");
        }
    }
}
