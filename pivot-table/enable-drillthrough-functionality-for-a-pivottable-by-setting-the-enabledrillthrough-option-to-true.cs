using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace Demo
{
    public class EnableDrillthroughDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
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

            // Add a pivot table covering the data range A1:B4, place it at D3, and name it "PivotTable1"
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Enable drill‑through (drilldown) functionality
            pivotTable.EnableDrilldown = true;

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Define output file path
            string outputPath = "EnableDrillthroughDemo.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
        }
    }
}