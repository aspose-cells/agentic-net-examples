using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using System.Drawing;

namespace AsposeCellsDemo
{
    public class ApplyStyleToPivotTable
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
                sheet.Cells["B1"].Value = "Amount";
                sheet.Cells["A2"].Value = "Fruit";
                sheet.Cells["B2"].Value = 100;
                sheet.Cells["A3"].Value = "Vegetable";
                sheet.Cells["B3"].Value = 150;
                sheet.Cells["A4"].Value = "Fruit";
                sheet.Cells["B4"].Value = 200;
                sheet.Cells["A5"].Value = "Vegetable";
                sheet.Cells["B5"].Value = 120;

                // Add a pivot table to the worksheet
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D2", "MyPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Define the fields for the pivot table
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Create a style that will be applied to the entire pivot table
                Style style = workbook.CreateStyle();
                style.Font.Name = "Calibri";
                style.Font.Size = 11;
                style.Font.IsBold = true;
                style.ForegroundColor = Color.LightGray;
                style.Pattern = BackgroundType.Solid;

                // Apply the style to all cells of the pivot table
                pivot.FormatAll(style);

                // Save the workbook with the formatted pivot table
                workbook.Save("PivotTable_Formatted.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ApplyStyleToPivotTable.Run();
        }
    }
}