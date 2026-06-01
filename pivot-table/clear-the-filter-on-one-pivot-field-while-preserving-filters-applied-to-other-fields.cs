using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class ClearSpecificPivotFieldFilter
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
                worksheet.Cells["A2"].Value = "Fruit";
                worksheet.Cells["A3"].Value = "Fruit";
                worksheet.Cells["A4"].Value = "Vegetable";
                worksheet.Cells["A5"].Value = "Vegetable";

                worksheet.Cells["B1"].Value = "Region";
                worksheet.Cells["B2"].Value = "North";
                worksheet.Cells["B3"].Value = "South";
                worksheet.Cells["B4"].Value = "North";
                worksheet.Cells["B5"].Value = "South";

                worksheet.Cells["C1"].Value = "Sales";
                worksheet.Cells["C2"].Value = 120;
                worksheet.Cells["C3"].Value = 150;
                worksheet.Cells["C4"].Value = 200;
                worksheet.Cells["C5"].Value = 180;

                // Add a pivot table based on the data range
                int pivotIndex = worksheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Apply filters on both row and column fields
                // Row field (Category) – filter out "Vegetable"
                pivotTable.PivotFilters.AddLabelFilter(0, PivotFilterType.CaptionNotEqual, "Vegetable", null);
                // Column field (Region) – filter out "South"
                pivotTable.PivotFilters.AddLabelFilter(1, PivotFilterType.CaptionNotEqual, "South", null);

                // Refresh and calculate to apply filters
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Clear filter only on the row field (Category, index 0) while keeping column filter intact.
                pivotTable.PivotFilters.ClearFilter(0);

                // Refresh again to reflect the change
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "ClearSpecificPivotFieldFilter.xlsx";
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
            ClearSpecificPivotFieldFilter.Run();
        }
    }
}