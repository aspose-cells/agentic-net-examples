using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class ChangePivotHeaderFontColor
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].Value = "Category";
                sheet.Cells["B1"].Value = "Value";
                sheet.Cells["A2"].Value = "A";
                sheet.Cells["B2"].Value = 10;
                sheet.Cells["A3"].Value = "B";
                sheet.Cells["B3"].Value = 20;
                sheet.Cells["A4"].Value = "A";
                sheet.Cells["B4"].Value = 30;
                sheet.Cells["A5"].Value = "B";
                sheet.Cells["B5"].Value = 40;

                // Add a pivot table at D3
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table (Category as row, Value as data)
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);
                pivotTable.CalculateData();

                // Create a style that sets the font color to black
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.Color = Color.Black; // Desired font color

                // Enable the font color flag so only the font color is applied
                StyleFlag flag = new StyleFlag { FontColor = true };

                // Apply the style to the pivot table header cell (top‑left cell of the pivot area)
                Cell headerCell = sheet.Cells["D3"]; // top‑left cell of the pivot table
                headerCell.SetStyle(headerStyle, flag);

                // Save the workbook
                workbook.Save("PivotHeaderFontColorBlack.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ChangePivotHeaderFontColor.Run();
        }
    }
}