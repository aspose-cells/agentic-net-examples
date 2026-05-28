using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelineFilterDemo
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data with a Date field and a numeric field
                cells["A1"].PutValue("Date");
                cells["B1"].PutValue("Sales");

                cells["A2"].PutValue(new DateTime(2023, 1, 5));
                cells["B2"].PutValue(1200);

                cells["A3"].PutValue(new DateTime(2023, 2, 12));
                cells["B3"].PutValue(1500);

                cells["A4"].PutValue(new DateTime(2023, 3, 20));
                cells["B4"].PutValue(1800);

                cells["A5"].PutValue(new DateTime(2023, 4, 8));
                cells["B5"].PutValue(2000);

                // Create a PivotTable based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add the Date field to the Row area and Sales to the Data area
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh and calculate the PivotTable
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a Timeline control linked to the Date field of the PivotTable.
                // Use row/column indices (E1 => row 0, column 4) to avoid cell‑name parsing issues.
                int timelineIndex = sheet.Timelines.Add(pivot, 0, 4, "Date");
                Timeline timeline = sheet.Timelines[timelineIndex];

                // Set a custom start date for the Timeline to focus on a specific period
                timeline.StartDate = new DateTime(2023, 2, 1);
                timeline.Caption = "Sales Timeline (Feb 2023 onward)";

                // Save the workbook with the Timeline filter applied
                workbook.Save("TimelineFilterDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}