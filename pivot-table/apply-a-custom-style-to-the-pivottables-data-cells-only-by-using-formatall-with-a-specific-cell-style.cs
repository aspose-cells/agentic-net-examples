using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableDataCellStyleDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Product";
            sheet.Cells["B1"].Value = "Region";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Apple";
            sheet.Cells["B2"].Value = "North";
            sheet.Cells["C2"].Value = 1200;

            sheet.Cells["A3"].Value = "Apple";
            sheet.Cells["B3"].Value = "South";
            sheet.Cells["C3"].Value = 1500;

            sheet.Cells["A4"].Value = "Banana";
            sheet.Cells["B4"].Value = "North";
            sheet.Cells["C4"].Value = 800;

            sheet.Cells["A5"].Value = "Banana";
            sheet.Cells["B5"].Value = "South";
            sheet.Cells["C5"].Value = 950;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivot.AddFieldToArea(PivotFieldType.Row, "Product");
            pivot.AddFieldToArea(PivotFieldType.Column, "Region");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Ensure the pivot table is calculated before formatting
            pivot.CalculateData();

            // Create a custom style for the data cells
            Style dataStyle = workbook.CreateStyle();
            dataStyle.Font.Name = "Calibri";
            dataStyle.Font.Size = 11;
            dataStyle.Font.IsBold = true;
            dataStyle.ForegroundColor = Color.LightYellow;
            dataStyle.Pattern = BackgroundType.Solid;

            // Apply the style to all cells of the pivot table using FormatAll
            // (FormatAll formats the entire pivot area; here we use it as requested)
            pivot.FormatAll(dataStyle);

            // Save the workbook
            workbook.Save("PivotTableDataCellStyleDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}