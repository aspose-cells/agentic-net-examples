using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class AddClearPivotFilterDemo
    {
        public static void Run()
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
            worksheet.Cells["B2"].Value = 100;
            worksheet.Cells["B3"].Value = 200;
            worksheet.Cells["B4"].Value = 150;
            worksheet.Cells["B5"].Value = 300;

            // Add a pivot table based on the data range
            int pivotIndex = worksheet.PivotTables.Add("A1:B5", "E3", "SalesPivot");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Add a label filter to keep only "Fruit"
            pivotTable.PivotFilters.AddLabelFilter(0, PivotFilterType.CaptionEqual, "Fruit", null);

            // Refresh and calculate to apply the filter
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Clear the filter
            pivotTable.PivotFilters.ClearFilter(0);

            // Refresh again to reflect the cleared filter
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("AddClearPivotFilterDemo.xlsx");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            AddClearPivotFilterDemo.Run();
        }
    }
}