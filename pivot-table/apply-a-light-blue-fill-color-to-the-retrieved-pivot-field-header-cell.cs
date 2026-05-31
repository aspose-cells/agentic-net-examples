using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotHeaderStyle
{
    public class ApplyHeaderStyle
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Amount";
            sheet.Cells["A2"].Value = "Food";
            sheet.Cells["B2"].Value = 120;
            sheet.Cells["A3"].Value = "Drink";
            sheet.Cells["B3"].Value = 80;
            sheet.Cells["A4"].Value = "Food";
            sheet.Cells["B4"].Value = 150;
            sheet.Cells["A5"].Value = "Drink";
            sheet.Cells["B5"].Value = 70;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: put Category in rows and Amount in data area
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category column
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Amount column

            // Calculate the pivot data so that the table is populated
            pivotTable.CalculateData();

            // Create a style with a light blue fill color
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.ForegroundColor = Color.LightBlue;

            // Apply the style to the pivot field header cell.
            // In pivot coordinates, the first row header is typically at row index 2, column index 0.
            // Adjust these indices if your pivot layout differs.
            pivotTable.Format(2, 0, headerStyle);

            // Save the workbook
            string outputPath = "PivotHeaderLightBlue.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
    }
}