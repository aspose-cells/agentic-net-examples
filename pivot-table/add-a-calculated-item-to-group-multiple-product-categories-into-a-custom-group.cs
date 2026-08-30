// Title: Create a calculated pivot item to group Apple and Banana into a custom "FruitGroup" using Aspose.Cells for C#
// AI Prompts: Generate a C# snippet that adds a calculated item named FruitGroup to a PivotTable row field with the formula =Apple+Banana using Aspose.Cells. | Write code to refresh the PivotTable, recalculate its data, and save the workbook as CustomGroupPivot.xlsx.
// Common Searches: Aspose.Cells C# how to add a calculated item to a pivot table row field | group specific rows in an Aspose.Cells pivot table programmatically | create custom product group in pivot table using Aspose.Cells API | calculate sum of selected items in an Aspose.Cells pivot table | refresh and calculate pivot data after adding a calculated item in C#
// Tags: add calculated item Aspose.Cells pivot | pivot table custom group C# | Aspose.Cells refresh pivot data | save workbook as xlsx Aspose.Cells | row field calculated item formula Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsCustomGroupDemo
{
    // The example creates a workbook, fills it with product and sales data, builds a pivot table, adds the Product field to the row area, defines a calculated item called "FruitGroup" that sums the Apple and Banana rows, refreshes and recalculates the pivot, and saves the file as CustomGroupPivot.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: Product category and Sales
            sheet.Cells["A1"].Value = "Product";
            sheet.Cells["B1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Apple";
            sheet.Cells["B2"].Value = 1200;

            sheet.Cells["A3"].Value = "Banana";
            sheet.Cells["B3"].Value = 800;

            sheet.Cells["A4"].Value = "Orange";
            sheet.Cells["B4"].Value = 1500;

            sheet.Cells["A5"].Value = "Grapes";
            sheet.Cells["B5"].Value = 900;

            // Create a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the Product field to the row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

            // Add the Sales field to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Get the row pivot field (Product)
            PivotField productField = pivotTable.RowFields[0];

            // Add a calculated item that groups Apple and Banana into a custom group named "FruitGroup"
            // The formula sums the two items; Excel syntax for pivot calculated items uses item names directly.
            productField.AddCalculatedItem("FruitGroup", "=Apple + Banana");

            // Refresh and calculate the pivot table to apply the new calculated item
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("CustomGroupPivot.xlsx");
        }
    }
}
