using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotTableDrilldownDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Pivot table created successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "SubCategory";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Electronics";
            sheet.Cells["B2"].Value = "Phones";
            sheet.Cells["C2"].Value = 1200;

            sheet.Cells["A3"].Value = "Electronics";
            sheet.Cells["B3"].Value = "Laptops";
            sheet.Cells["C3"].Value = 2500;

            sheet.Cells["A4"].Value = "Furniture";
            sheet.Cells["B4"].Value = "Chairs";
            sheet.Cells["C4"].Value = 800;

            sheet.Cells["A5"].Value = "Furniture";
            sheet.Cells["B5"].Value = "Tables";
            sheet.Cells["C5"].Value = 1500;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");          // Row field
            pivot.AddFieldToArea(PivotFieldType.Column, "SubCategory");   // Column field
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");           // Data field

            // Enable drill‑down functionality and show the expand/collapse buttons
            pivot.EnableDrilldown = true; // Allows double‑click to see underlying records
            pivot.ShowDrill = true;       // Displays drill indicators in the UI
            pivot.PrintDrill = true;      // Prints indicators when the sheet is printed

            // Refresh and calculate the pivot table data
            pivot.RefreshData();
            pivot.CalculateData();

            // Define output file path
            string outputPath = "PivotTableDrilldownDemo.xlsx";

            // Save the workbook if the directory is writable
            try
            {
                workbook.Save(outputPath);
            }
            catch (Exception saveEx)
            {
                Console.Error.WriteLine($"Failed to save workbook: {saveEx.Message}");
                throw;
            }
        }
    }
}