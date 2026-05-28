using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class RefreshPivotAfterHeaderFormatting
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Amount");
                worksheet.Cells["A2"].PutValue("Fruit");
                worksheet.Cells["B2"].PutValue(120);
                worksheet.Cells["A3"].PutValue("Vegetable");
                worksheet.Cells["B3"].PutValue(80);
                worksheet.Cells["A4"].PutValue("Fruit");
                worksheet.Cells["B4"].PutValue(150);
                worksheet.Cells["A5"].PutValue("Vegetable");
                worksheet.Cells["B5"].PutValue(70);

                // Add a pivot table based on the data range
                int pivotIndex = worksheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Populate the pivot table
                pivotTable.CalculateData();

                // Preserve formatting when the pivot table is refreshed
                pivotTable.PreserveFormatting = true;

                // Create a style for the header cell (bold white font on dark blue background)
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.IsBold = true;
                headerStyle.Font.Color = Color.White;
                headerStyle.ForegroundColor = Color.DarkBlue;
                headerStyle.Pattern = BackgroundType.Solid;

                // Apply the style to the "Amount" column header in the pivot table
                // Pivot table coordinates: row 1 = column headers, column 1 = first data field
                pivotTable.Format(1, 1, headerStyle);

                // Refresh the pivot table (preserves the header style)
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Define output file path
                string outputPath = "PivotTable_RefreshAfterHeaderFormatting.xlsx";

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RefreshPivotAfterHeaderFormatting.Run();
        }
    }
}