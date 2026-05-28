using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Drawing;

namespace PivotHeaderThemeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Region";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Fruit";
            sheet.Cells["B2"].Value = "North";
            sheet.Cells["C2"].Value = 1200;

            sheet.Cells["A3"].Value = "Fruit";
            sheet.Cells["B3"].Value = "South";
            sheet.Cells["C3"].Value = 850;

            sheet.Cells["A4"].Value = "Vegetable";
            sheet.Cells["B4"].Value = "North";
            sheet.Cells["C4"].Value = 640;

            sheet.Cells["A5"].Value = "Vegetable";
            sheet.Cells["B5"].Value = "South";
            sheet.Cells["C5"].Value = 730;

            // Add a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];

            // Configure the pivot fields
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Column, "Region");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Ensure the pivot table is calculated
            pivot.CalculateData();

            // Define a custom theme color (Accent1) – for example, a teal shade
            Color customAccent = Color.FromArgb(0, 128, 128);
            workbook.SetThemeColor(ThemeColorType.Accent1, customAccent);

            // Create a style that uses the theme color for the font
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0);
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Size = 12;
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.ForegroundColor = Color.LightGray; // optional background for visibility

            // Apply the style to the entire pivot table using FormatAll
            // (Headers will inherit the font theme color defined above)
            pivot.FormatAll(headerStyle);

            // Save the workbook
            workbook.Save("PivotHeaderThemeDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}