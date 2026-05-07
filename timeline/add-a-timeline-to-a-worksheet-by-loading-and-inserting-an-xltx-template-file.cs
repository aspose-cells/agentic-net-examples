using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsTimelineExample
{
    class Program
    {
        static void Main()
        {
            // Load the XLTX template file (the template should contain at least a date column named "Date")
            Workbook workbook = new Workbook("Template.xltx");

            // Get the first worksheet from the loaded workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Create a PivotTable that will serve as the data source for the Timeline.
            int pivotIndex = sheet.PivotTables.Add("A1:A4", "C1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add the "Date" field to the Row area of the PivotTable
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");

            // Refresh and calculate the PivotTable so that it contains valid data
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook
            workbook.Save("TimelineResult.xlsx");
        }
    }
}