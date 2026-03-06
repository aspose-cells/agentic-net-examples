using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using System.Drawing;

namespace AsposeCellsPivotStyleDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Product";
            sheet.Cells["B1"].Value = "Region";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Laptop";
            sheet.Cells["B2"].Value = "North";
            sheet.Cells["C2"].Value = 1200;

            sheet.Cells["A3"].Value = "Desktop";
            sheet.Cells["B3"].Value = "South";
            sheet.Cells["C3"].Value = 800;

            sheet.Cells["A4"].Value = "Tablet";
            sheet.Cells["B4"].Value = "East";
            sheet.Cells["C4"].Value = 500;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Create a style to be applied to the entire pivot table
            Style style = workbook.CreateStyle();
            style.Font.Name = "Calibri";
            style.Font.Size = 11;
            style.Font.IsBold = true;
            style.ForegroundColor = Color.LightGray;
            style.Pattern = BackgroundType.Solid;

            // Apply the style to all cells of the pivot table
            pivotTable.FormatAll(style);

            // Save the workbook in XLSX format
            workbook.Save("PivotTableStyleAllDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}