using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using System.Drawing;

namespace AsposeCellsPivotFormatAllDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Year";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Electronics";
            sheet.Cells["B2"].Value = 2020;
            sheet.Cells["C2"].Value = 1500;

            sheet.Cells["A3"].Value = "Electronics";
            sheet.Cells["B3"].Value = 2021;
            sheet.Cells["C3"].Value = 1800;

            sheet.Cells["A4"].Value = "Furniture";
            sheet.Cells["B4"].Value = 2020;
            sheet.Cells["C4"].Value = 1200;

            sheet.Cells["A5"].Value = "Furniture";
            sheet.Cells["B5"].Value = 2021;
            sheet.Cells["C5"].Value = 1300;

            // Add a pivot table based on the data range
            PivotTableCollection pivotTables = sheet.PivotTables;
            int pivotIndex = pivotTables.Add("=Sheet1!A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = pivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Year");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Create a style that will be applied to the entire pivot table
            Style style = workbook.CreateStyle();
            style.Font.Name = "Calibri";
            style.Font.Size = 11;
            style.Font.IsBold = true;
            style.ForegroundColor = Color.LightYellow;
            style.Pattern = BackgroundType.Solid;

            // Apply the style to all cells of the pivot table
            pivotTable.FormatAll(style);

            // Save the workbook
            workbook.Save("PivotTableFormatAllDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}