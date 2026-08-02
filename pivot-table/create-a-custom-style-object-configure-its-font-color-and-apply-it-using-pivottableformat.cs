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
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Amount";
            sheet.Cells["A2"].Value = "Food";
            sheet.Cells["B2"].Value = 120;
            sheet.Cells["A3"].Value = "Food";
            sheet.Cells["B3"].Value = 80;
            sheet.Cells["A4"].Value = "Drink";
            sheet.Cells["B4"].Value = 150;
            sheet.Cells["A5"].Value = "Drink";
            sheet.Cells["B5"].Value = 70;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Amount as data field

            // Calculate the pivot table data
            pivotTable.CalculateData();

            // Create a custom style and set its font color
            Style customStyle = workbook.CreateStyle();
            customStyle.Font.Color = Color.DarkRed; // Desired font color
            customStyle.Font.IsBold = true;         // Optional: make the font bold

            // Apply the style to the data body range of the pivot table using Format
            CellArea dataArea = pivotTable.DataBodyRange;
            pivotTable.Format(dataArea, customStyle);

            // Save the workbook
            workbook.Save("PivotTableCustomFontColor.xlsx");
        }
    }
}