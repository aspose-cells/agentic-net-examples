using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExpandAllDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(200);
                sheet.Cells["A4"].PutValue("A");
                sheet.Cells["B4"].PutValue(150);

                // Add a pivot table to the worksheet
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

                // Enable drilldown and show drill buttons (required for expand/collapse)
                pivotTable.EnableDrilldown = true;
                pivotTable.ShowDrill = true;

                // Refresh data and calculate the pivot table
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // NOTE: Aspose.Cells does not provide a direct ExpandAll method.
                // The pivot table will display all data when refreshed with drill buttons enabled.

                // Save the workbook
                workbook.Save("ExpandAllPivotDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}