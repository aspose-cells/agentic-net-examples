using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotTableCustomFontColorDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                worksheet.Cells["A1"].Value = "Category";
                worksheet.Cells["B1"].Value = "Amount";
                worksheet.Cells["A2"].Value = "A";
                worksheet.Cells["B2"].Value = 100;
                worksheet.Cells["A3"].Value = "B";
                worksheet.Cells["B3"].Value = 200;

                // Add a pivot table based on the data range
                int pivotIndex = worksheet.PivotTables.Add("A1:B3", "D5", "PivotTable1");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Amount as data field
                pivotTable.CalculateData(); // Populate the pivot table

                // Create a custom style and set its font color
                Style customStyle = workbook.CreateStyle();
                customStyle.Font.Color = Color.Red; // Desired font color

                // Apply the custom style to a specific cell in the pivot table
                // Row and column indices are relative to the pivot table's own coordinate system
                // Here we format the header cell (row 2, column 0) as an example
                pivotTable.Format(2, 0, customStyle);

                // Save the workbook with the applied style
                string outputPath = "PivotTableCustomFontColor.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            PivotTableCustomFontColorDemo.Run();
        }
    }
}