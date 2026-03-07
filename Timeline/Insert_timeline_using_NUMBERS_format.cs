using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelineNumbersFormat
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare sample data in a DataTable
            DataTable table = new DataTable();
            table.Columns.Add("Date", typeof(DateTime));
            table.Columns.Add("Value", typeof(double));

            table.Rows.Add(new DateTime(2023, 1, 1), 100);
            table.Rows.Add(new DateTime(2023, 2, 1), 200);
            table.Rows.Add(new DateTime(2023, 3, 1), 300);
            table.Rows.Add(new DateTime(2023, 4, 1), 400);

            // Import the data into the worksheet
            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = true
            };
            sheet.Cells.ImportData(table, 0, 0, importOptions);

            // Create a pivot table that uses the imported data as its source
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add the Date field to the Row area (required for Timeline)
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            // Add the Value field to the Data area
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // Refresh and calculate the pivot table
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline control linked to the pivot table
            int timelineIndex = sheet.Timelines.Add(pivot, "E1", "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];
            timeline.Caption = "Date Timeline (Numbers Format)";

            // Save the workbook
            workbook.Save("TimelineWithNumbersFormat.xlsx");
        }
    }
}