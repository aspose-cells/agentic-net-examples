using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class DisableRowGrandTotalsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].Value = "Product";
                sheet.Cells["B1"].Value = "Region";
                sheet.Cells["C1"].Value = "Sales";

                sheet.Cells["A2"].Value = "Product A";
                sheet.Cells["B2"].Value = "North";
                sheet.Cells["C2"].Value = 1000;

                sheet.Cells["A3"].Value = "Product B";
                sheet.Cells["B3"].Value = "South";
                sheet.Cells["C3"].Value = 2000;

                sheet.Cells["A4"].Value = "Product C";
                sheet.Cells["B4"].Value = "East";
                sheet.Cells["C4"].Value = 3000;

                sheet.Cells["A5"].Value = "Product A";
                sheet.Cells["B5"].Value = "West";
                sheet.Cells["C5"].Value = 1500;

                sheet.Cells["A6"].Value = "Product B";
                sheet.Cells["B6"].Value = "North";
                sheet.Cells["C6"].Value = 2500;

                // Add a pivot table based on the data range
                PivotTableCollection pivotTables = sheet.PivotTables;
                int pivotIndex = pivotTables.Add("A1:C6", "E3", "PivotTable1");
                PivotTable pivotTable = pivotTables[pivotIndex];

                // Configure the pivot table: add row and data fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Disable row grand totals to simplify the summary view
                pivotTable.ShowRowGrandTotals = false;

                // Recalculate the pivot table after changing the setting
                pivotTable.CalculateData();

                // Define output file path
                string outputPath = "DisableRowGrandTotalsDemo.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point required for console application
    public class Program
    {
        public static void Main(string[] args)
        {
            DisableRowGrandTotalsDemo.Run();
        }
    }
}