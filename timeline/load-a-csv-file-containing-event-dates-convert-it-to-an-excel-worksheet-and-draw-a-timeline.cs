using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class TimelineFromCsv
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Import the CSV file into the worksheet
            string csvPath = "events.csv";          // CSV file containing Event and Date columns
            string delimiter = ",";                // CSV delimiter
            bool convertNumeric = true;            // Convert numeric strings to numbers
            int startRow = 0;                      // Start importing at cell A1 (row index 0)
            int startColumn = 0;                   // Start importing at column A (column index 0)

            // Prevent FileNotFoundException
            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"CSV file not found: {Path.GetFullPath(csvPath)}");
                return;
            }

            cells.ImportCSV(csvPath, delimiter, convertNumeric, startRow, startColumn);

            // 3. Create a PivotTable that will serve as the data source for the Timeline
            // Assuming the CSV has two columns: "Event" (A) and "Date" (B)
            // Use a range that covers the imported data (e.g., A1:B100)
            int pivotIndex = sheet.PivotTables.Add("A1:B100", "D1", "EventPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add the Date field as the row axis (required for a Timeline)
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");

            // Add the Event field as a data field (count of events per date)
            pivot.AddFieldToArea(PivotFieldType.Data, "Event");

            // Refresh the PivotTable to populate it with data
            pivot.RefreshData();
            pivot.CalculateData();

            // 4. Add a Timeline control linked to the PivotTable's Date field
            // Place the Timeline starting at cell G5
            int timelineIndex = sheet.Timelines.Add(pivot, "G5", "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];

            // Optional: customize the Timeline appearance via its Shape property
            timeline.Shape.Top = 200;      // vertical offset in pixels
            timeline.Shape.Left = 50;      // horizontal offset in pixels
            timeline.Shape.Width = 600;    // width in pixels
            timeline.Shape.Height = 80;    // height in pixels
            timeline.Caption = "Event Timeline";

            // 5. Save the workbook with the Timeline embedded
            string outputPath = "TimelineFromCsv.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}