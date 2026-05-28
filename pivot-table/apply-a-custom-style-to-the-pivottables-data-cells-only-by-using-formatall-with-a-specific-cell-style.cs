using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using System.Drawing;

namespace PivotTableCustomStyleDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            worksheet.Cells["A1"].Value = "Category";
            worksheet.Cells["B1"].Value = "Amount";
            worksheet.Cells["A2"].Value = "Food";
            worksheet.Cells["B2"].Value = 120;
            worksheet.Cells["A3"].Value = "Food";
            worksheet.Cells["B3"].Value = 80;
            worksheet.Cells["A4"].Value = "Drink";
            worksheet.Cells["B4"].Value = 150;
            worksheet.Cells["A5"].Value = "Drink";
            worksheet.Cells["B5"].Value = 70;

            // Add a pivot table based on the data range
            PivotTableCollection pivotTables = worksheet.PivotTables;
            int pivotIndex = pivotTables.Add("A1:B5", "D3", "SalesPivot");
            PivotTable pivotTable = pivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Create a custom style for the data cells
            Style dataStyle = workbook.CreateStyle();
            dataStyle.Font.Name = "Calibri";
            dataStyle.Font.Size = 11;
            dataStyle.Font.IsBold = true;
            dataStyle.ForegroundColor = Color.LightYellow;
            dataStyle.Pattern = BackgroundType.Solid;

            // Apply the custom style to all cells in the pivot table (including data cells)
            // Since FormatAll applies to the whole pivot area, we rely on PreserveFormatting
            // to keep the style when the pivot is refreshed.
            pivotTable.PreserveFormatting = true;
            pivotTable.FormatAll(dataStyle);

            // Save the workbook
            workbook.Save("PivotTableDataStyleDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}