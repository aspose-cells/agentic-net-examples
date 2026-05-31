using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable related classes

namespace PivotLabelFilterDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Alpha");
                sheet.Cells["A3"].PutValue("Beta");
                sheet.Cells["A4"].PutValue("Alpha");
                sheet.Cells["A5"].PutValue("Gamma");

                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["B3"].PutValue(200);
                sheet.Cells["B4"].PutValue(150);
                sheet.Cells["B5"].PutValue(300);

                // Create a pivot table covering the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add the "Category" field to the row area and "Amount" to the data area
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Retrieve the row field and apply a label filter to include only "Alpha"
                PivotField rowField = pivot.RowFields[0];
                rowField.FilterByLabel(PivotFilterType.CaptionEqual, "Alpha", null);

                // Refresh the pivot table to apply the filter and recalculate data
                pivot.RefreshData();
                pivot.CalculateData();

                // Save the workbook
                string outputPath = "PivotLabelFilterDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}