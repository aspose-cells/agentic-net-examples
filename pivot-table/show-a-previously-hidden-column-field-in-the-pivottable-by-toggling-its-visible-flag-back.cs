using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    public class ShowHiddenColumnField
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data for the pivot table
                cells["A1"].Value = "Category";
                cells["B1"].Value = "Region";
                cells["C1"].Value = "Sales";

                cells["A2"].Value = "Food";
                cells["B2"].Value = "North";
                cells["C2"].Value = 1200;

                cells["A3"].Value = "Food";
                cells["B3"].Value = "South";
                cells["C3"].Value = 1500;

                cells["A4"].Value = "Beverage";
                cells["B4"].Value = "North";
                cells["C4"].Value = 800;

                cells["A5"].Value = "Beverage";
                cells["B5"].Value = "South";
                cells["C5"].Value = 950;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields: Category as row, Region as column, Sales as data
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Hide the column field "Region"
                PivotField columnField = pivotTable.ColumnFields[0];
                columnField.ShowAllItems = false;

                // Refresh and calculate to apply the hide state
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Show the previously hidden column field
                columnField.ShowAllItems = true;

                // Refresh and calculate again to reflect the change
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "PivotTable_ShowHiddenColumnField.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
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
            ShowHiddenColumnField.Run();
        }
    }
}