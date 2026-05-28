using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotFieldCombinedFiltersDemo
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
                // Column A: Category, Column B: Sales
                cells["A1"].Value = "Category";
                cells["A2"].Value = "Apple";
                cells["A3"].Value = "Banana";
                cells["A4"].Value = "Apple";
                cells["A5"].Value = "Cherry";
                cells["A6"].Value = "Banana";
                cells["A7"].Value = "Apple";

                cells["B1"].Value = "Sales";
                cells["B2"].Value = 120;
                cells["B3"].Value = 80;
                cells["B4"].Value = 150;
                cells["B5"].Value = 200;
                cells["B6"].Value = 90;
                cells["B7"].Value = 130;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B7", "D3", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add the Category field as a row field and Sales as a data field
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Enable multiple filters on the same pivot field
                pivot.AllowMultipleFiltersPerField = true;

                // Get the row field (Category) to apply filters
                PivotField categoryField = pivot.RowFields[0];

                // 1) Apply a Top 10 filter: show top 2 categories by count
                // valueFieldIndex = 0 because there is only one data field (Sales) at index 0 in the data region
                categoryField.FilterTop10(
                    valueFieldIndex: 0,
                    type: PivotFilterType.Count,
                    isTop: true,
                    itemCount: 2);

                // 2) Apply a label filter: show categories that begin with "A"
                categoryField.FilterByLabel(
                    type: PivotFilterType.CaptionBeginsWith,
                    label1: "A",
                    label2: null);

                // Refresh and calculate the pivot table to reflect the filters
                pivot.RefreshData();
                pivot.CalculateData();

                // Save the workbook
                string outputPath = "PivotFieldCombinedFiltersDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                PivotFieldCombinedFiltersDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}