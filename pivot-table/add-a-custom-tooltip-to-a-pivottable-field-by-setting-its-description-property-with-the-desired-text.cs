using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotTableCustomTooltipDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Region");
                worksheet.Cells["C1"].PutValue("Sales");

                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue("North");
                worksheet.Cells["C2"].PutValue(1200);

                worksheet.Cells["A3"].PutValue("Apple");
                worksheet.Cells["B3"].PutValue("South");
                worksheet.Cells["C3"].PutValue(800);

                worksheet.Cells["A4"].PutValue("Banana");
                worksheet.Cells["B4"].PutValue("North");
                worksheet.Cells["C4"].PutValue(1500);

                worksheet.Cells["A5"].PutValue("Banana");
                worksheet.Cells["B5"].PutValue("South");
                worksheet.Cells["C5"].PutValue(1100);

                // Add a pivot table to the worksheet
                PivotTableCollection pivotTables = worksheet.PivotTables;
                int pivotIndex = pivotTables.Add("A1:C5", "E3", "SalesPivot");
                PivotTable pivotTable = pivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Set a custom tooltip (alt text description) for the pivot table
                pivotTable.AltTextDescription = "This pivot table shows sales distribution by product and region.";
                pivotTable.AltTextTitle = "Sales Pivot Table Tooltip";

                // Refresh and calculate the pivot table data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Ensure the output file can be written (delete if it already exists)
                string outputPath = "PivotTableCustomTooltipDemo.xlsx";
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            PivotTableCustomTooltipDemo.Run();
        }
    }
}