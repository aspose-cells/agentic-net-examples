// Title: Hide a PivotField in an Aspose.Cells PivotTable with C#
// Description: Creates a workbook, populates sample sales data, builds a PivotTable, adds Category (row), Product (column) and Sales (data) fields, then hides the Product field by removing it from the column area, refreshes the pivot, and saves the file.
// Keywords: Aspose.Cells hide pivot field | C# PivotTable remove column field | Aspose.Cells PivotField Visible false | hide column field Aspose.Cells .NET | remove pivot field programmatically
// Common Searches: how to hide a pivot column field in Aspose.Cells C# | remove pivot field from report area Aspose.Cells | set PivotField visibility false Aspose.Cells | Aspose.Cells hide product field in pivot table | C# hide specific PivotField Aspose.Cells
// Developer Intent: Programmatically hide the "Product" PivotField from the column area of an Aspose.Cells PivotTable using C#.
// Use Cases: Create cleaner reports by omitting unnecessary column fields. | Allow end‑users to toggle pivot fields on or off in a .NET application. | Prepare workbooks for export where certain pivot fields must remain hidden.
// AI Prompts: Write C# code that hides a specific PivotField in an Aspose.Cells PivotTable without deleting the field definition. | Show how to toggle the Visible property of a PivotField and refresh the pivot in Aspose.Cells. | Explain the steps to remove a pivot field from its area, recalculate the PivotTable, and save the workbook using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotFieldHideDemo
{
    // Creates a workbook, populates sample sales data, builds a PivotTable, adds Category (row), Product (column) and Sales (data) fields, then hides the Product field by removing it from the column area, refreshes the pivot, and saves the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Product";
            cells["C1"].Value = "Sales";

            cells["A2"].Value = "Fruit";
            cells["B2"].Value = "Apple";
            cells["C2"].Value = 1200;

            cells["A3"].Value = "Fruit";
            cells["B3"].Value = "Banana";
            cells["C3"].Value = 800;

            cells["A4"].Value = "Vegetable";
            cells["B4"].Value = "Carrot";
            cells["C4"].Value = 600;

            cells["A5"].Value = "Vegetable";
            cells["B5"].Value = "Tomato";
            cells["C5"].Value = 900;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            // Row field: Category
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            // Column field: Product
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
            // Data field: Sales
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // ------------------------------------------------------------
            // Hide the "Product" pivot field from the report area.
            // This is achieved by removing the field from its current area.
            // ------------------------------------------------------------
            pivotTable.RemoveField(PivotFieldType.Column, "Product");

            // Refresh and calculate the pivot table after modification
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook with the hidden pivot field
            workbook.Save("PivotFieldHiddenDemo.xlsx");
        }
    }
}
