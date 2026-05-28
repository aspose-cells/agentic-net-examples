using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class ClearPivotRowFieldFilterDemo
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
                worksheet.Cells["A3"].Value = "Vegetable";
                worksheet.Cells["A4"].Value = "Fruit";
                worksheet.Cells["A5"].Value = "Vegetable";

                worksheet.Cells["B1"].Value = "Sales";
                worksheet.Cells["B2"].Value = 120;
                worksheet.Cells["B3"].Value = 80;
                worksheet.Cells["B4"].Value = 150;
                worksheet.Cells["B5"].Value = 200;

                // Add a pivot table covering the data range
                int pivotIndex = worksheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

                // Add the "Category" column as a row field
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);

                // Add the "Sales" column as a data field
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

                // Apply a filter on the row field to show only "Fruit"
                PivotField rowField = pivotTable.RowFields[0];
                rowField.FilterByLabel(PivotFilterType.CaptionEqual, "Fruit", null);

                // Clear the filter on the row field, restoring all items
                rowField.ClearFilter();

                // Refresh the pivot table to reflect the cleared filter
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "ClearPivotRowFieldFilterDemo.xlsx";
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
            ClearPivotRowFieldFilterDemo.Run();
        }
    }
}