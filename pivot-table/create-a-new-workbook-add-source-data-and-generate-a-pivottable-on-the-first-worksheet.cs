using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    public class CreatePivotTable
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate source data
                Cells cells = worksheet.Cells;
                cells["A1"].Value = "Product";
                cells["B1"].Value = "Region";
                cells["C1"].Value = "Sales";

                cells["A2"].Value = "Product1";
                cells["B2"].Value = "North";
                cells["C2"].Value = 1000;

                cells["A3"].Value = "Product2";
                cells["B3"].Value = "South";
                cells["C3"].Value = 2000;

                cells["A4"].Value = "Product3";
                cells["B4"].Value = "East";
                cells["C4"].Value = 3000;

                cells["A5"].Value = "Product1";
                cells["B5"].Value = "West";
                cells["C5"].Value = 4000;

                cells["A6"].Value = "Product2";
                cells["B6"].Value = "North";
                cells["C6"].Value = 5000;

                // Define source range, destination cell and table name
                string sourceData = "A1:C6";
                string destCellName = "E2";
                string tableName = "SalesPivotTable";

                // Add a new PivotTable
                int pivotIndex = worksheet.PivotTables.Add(sourceData, destCellName, tableName);
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

                // Configure pivot fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh and calculate the pivot data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "PivotTableDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            CreatePivotTable.Run();
        }
    }
}