using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace TimelineDemo
{
    class TimelineFromCsv
    {
        static void Main()
        {
            try
            {
                // Path to the source CSV file containing event dates
                string csvPath = "events.csv";

                // Verify that the CSV file exists before attempting to import
                if (!File.Exists(csvPath))
                {
                    Console.WriteLine($"CSV file not found: {Path.GetFullPath(csvPath)}");
                    return;
                }

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Import the CSV data into the worksheet starting at cell A1
                cells.ImportCSV(csvPath, ",", true, 0, 0);

                // Define the source range for the PivotTable (adjust as needed)
                string sourceRange = "A1:C100";

                // Add a PivotTable that will serve as the data source for the Timeline
                int pivotIndex = sheet.PivotTables.Add(sourceRange, "E1", "EventPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add the date field to the Row area (must exist in the CSV header as "Date")
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                // Optionally add other fields (e.g., Event name) to the Row area
                // pivot.AddFieldToArea(PivotFieldType.Row, "Event");

                // Refresh and calculate the PivotTable data
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a Timeline linked to the PivotTable using the "Date" field
                int timelineIndex = sheet.Timelines.Add(pivot, 0, 0, "Date");
                Timeline timeline = sheet.Timelines[timelineIndex];

                // Customize Timeline appearance via its Shape (optional)
                timeline.Shape.Top = 20;      // vertical offset in pixels
                timeline.Shape.Left = 20;     // horizontal offset in pixels
                timeline.Shape.Width = 500;   // width in pixels
                timeline.Shape.Height = 100;  // height in pixels
                timeline.Caption = "Event Timeline";

                // Save the workbook
                workbook.Save("EventTimeline.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved as EventTimeline.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}