using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotTableDisplayItemLabelsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B3", "D1", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table (row and data)
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Value as data field

                // NOTE: The DisplayItemLabels property is not available in the current Aspose.Cells API.
                // The pivot table will display default labels without this setting.

                // Refresh and calculate the pivot table data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook with the configured pivot table
                string outputPath = "PivotTableDisplayItemLabelsDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    internal class Program
    {
        private static void Main(string[] args)
        {
            PivotTableDisplayItemLabelsDemo.Run();
        }
    }
}