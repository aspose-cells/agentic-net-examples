using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class DisplayNullAsNA
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data with null values
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(null);   // Null sales
                sheet.Cells["A4"].PutValue(null);   // Null product
                sheet.Cells["B4"].PutValue(80);

                // Add a pivot table covering the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Configure the pivot table: product as row, sales as data
                pivot.AddFieldToArea(PivotFieldType.Row, 0);
                pivot.AddFieldToArea(PivotFieldType.Data, 1);

                // Enable custom display for null values and set the string to "N/A"
                pivot.DisplayNullString = true;
                pivot.NullString = "N/A";

                // Refresh and calculate the pivot table to apply changes
                pivot.RefreshData();
                pivot.CalculateData();

                // Save the workbook
                string outputPath = "PivotTable_NA_Display.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
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
            DisplayNullAsNA.Run();
        }
    }
}