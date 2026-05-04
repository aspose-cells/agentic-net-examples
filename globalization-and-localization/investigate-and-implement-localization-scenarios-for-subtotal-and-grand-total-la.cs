using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotLocalizationDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Electronics");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["A3"].PutValue("Furniture");
            sheet.Cells["B3"].PutValue(800);
            sheet.Cells["A4"].PutValue("Electronics");
            sheet.Cells["B4"].PutValue(500);
            sheet.Cells["A5"].PutValue("Furniture");
            sheet.Cells["B5"].PutValue(700);

            // Create a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:B5", "D1", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
            int dataFieldIdx = pivot.AddFieldToArea(PivotFieldType.Data, 1); // Sales as data field
            // Set the aggregation function for the data field (Sum)
            pivot.DataFields[dataFieldIdx].Function = ConsolidationFunction.Sum;

            // Refresh and calculate the pivot table
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook
            workbook.Save("PivotLocalizationDemo.xlsx");
        }

        public static void Main(string[] args)
        {
            Run();
        }
    }
}