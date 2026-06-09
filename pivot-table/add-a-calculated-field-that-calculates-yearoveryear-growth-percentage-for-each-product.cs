using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotTableYoYGrowthDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data: Product, Year, Sales
                cells["A1"].Value = "Product";
                cells["B1"].Value = "Year";
                cells["C1"].Value = "Sales";

                // Sample rows
                cells["A2"].Value = "Bike";   cells["B2"].Value = 2020; cells["C2"].Value = 1000;
                cells["A3"].Value = "Bike";   cells["B3"].Value = 2021; cells["C3"].Value = 1200;
                cells["A4"].Value = "Car";    cells["B4"].Value = 2020; cells["C4"].Value = 2000;
                cells["A5"].Value = "Car";    cells["B5"].Value = 2021; cells["C5"].Value = 2500;
                cells["A6"].Value = "Truck";  cells["B6"].Value = 2020; cells["C6"].Value = 1500;
                cells["A7"].Value = "Truck";  cells["B7"].Value = 2021; cells["C7"].Value = 1800;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C7", "E3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields: Product to rows, Year to columns, Sales to data area
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Year");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Add a calculated field that references Sales.
                // The YoY calculation will be handled by ShowValuesAs.
                pivotTable.AddCalculatedField("YoYGrowth", "=Sales", true);

                // Retrieve the newly added calculated field (last data field)
                PivotField yoyField = pivotTable.DataFields[pivotTable.DataFields.Count - 1];

                // Configure the calculated field to show percentage difference from the previous year
                PivotField yearColumnField = pivotTable.ColumnFields[0];
                yoyField.ShowValuesAs(
                    PivotFieldDataDisplayFormat.PercentageDifferenceFrom,
                    yearColumnField.BaseIndex,
                    PivotItemPositionType.Previous,
                    0); // baseItem is ignored when using Previous

                // Refresh and calculate the pivot table
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook (ensure the directory exists)
                string outputPath = "PivotTableYoYGrowth.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
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
            PivotTableYoYGrowthDemo.Run();
        }
    }
}