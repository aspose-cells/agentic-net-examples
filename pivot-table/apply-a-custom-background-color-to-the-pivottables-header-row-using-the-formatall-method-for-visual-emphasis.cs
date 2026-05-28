using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using System.Drawing;

namespace AsposeCellsPivotHeaderFormatting
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
            sheet.Cells["B1"].Value = "Amount";
            sheet.Cells["A2"].Value = "Fruit";
            sheet.Cells["B2"].Value = 120;
            sheet.Cells["A3"].Value = "Vegetable";
            sheet.Cells["B3"].Value = 80;
            sheet.Cells["A4"].Value = "Fruit";
            sheet.Cells["B4"].Value = 150;
            sheet.Cells["A5"].Value = "Vegetable";
            sheet.Cells["B5"].Value = 60;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: rows and data fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Calculate the pivot table data
            pivotTable.CalculateData();

            // Create a style with a custom background color for emphasis
            Style headerStyle = workbook.CreateStyle();
            headerStyle.ForegroundColor = Color.LightCoral;   // Desired background color
            headerStyle.Pattern = BackgroundType.Solid;      // Apply solid fill
            headerStyle.Font.IsBold = true;                  // Optional: make header text bold

            // Apply the style to the entire pivot table using FormatAll
            // (This highlights the header row as part of the whole table)
            pivotTable.FormatAll(headerStyle);

            // Save the workbook
            workbook.Save("PivotTableHeaderFormatted.xlsx", SaveFormat.Xlsx);
        }
    }
}