// Title: C# – Build an Excel Timeline from CSV Event Dates with Aspose.Cells
// Description: The sample checks for an events CSV, creates a workbook, imports the data, generates a PivotTable using the Date column as the row field and Event as the data field, adds a linked Timeline control, sets a caption, and saves the file as an interactive Excel timeline.
// Keywords: Aspose.Cells | C# CSV import | Excel timeline control | PivotTable Aspose.Cells | event timeline Excel | Aspose.Cells Timeline API | ImportCSV Aspose.Cells | AddTimeline Aspose.Cells | generate timeline from CSV
// Common Searches: Aspose.Cells import CSV and create timeline | Add timeline control to PivotTable using C# | Generate Excel timeline from event dates | Aspose.Cells timeline example .NET | How to link a timeline to a pivot table in Aspose.Cells
// Developer Intent: Create an Excel workbook that visualizes event dates with an interactive Timeline control sourced from a CSV file.
// Use Cases: Import a CSV containing Event and Date columns into a worksheet. | Build a PivotTable from the imported range, placing Date in the row area and Event in the data area. | Attach a Timeline control to the PivotTable, position it on the sheet, and customize its caption. | Save the workbook as an .xlsx file that end‑users can filter via the timeline.
// AI Prompts: Show C# code to customize the Timeline control's date range, style, and slicer layout with Aspose.Cells. | Explain how to add multiple Timeline controls for different date fields in the same workbook using Aspose.Cells for .NET. | Provide a step‑by‑step guide to export the generated timeline workbook to PDF while preserving the interactive elements.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

// The sample checks for an events CSV, creates a workbook, imports the data, generates a PivotTable using the Date column as the row field and Event as the data field, adds a linked Timeline control, sets a caption, and saves the file as an interactive Excel timeline.
class EventTimelineGenerator
{
    static void Main()
    {
        try
        {
            // Ensure the CSV file exists; create a simple sample if missing
            string csvPath = "events.csv";
            if (!File.Exists(csvPath))
            {
                var sample = "Event,Date\r\nLaunch,2023-01-15\r\nUpdate,2023-03-10\r\nRelease,2023-06-05\r\n";
                File.WriteAllText(csvPath, sample);
            }

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Import the CSV file (headers "Event,Date")
            cells.ImportCSV(csvPath, ",", true, 0, 0);

            // Determine the used range for the pivot source
            int lastRow = cells.MaxDataRow + 1; // +1 because range is inclusive
            string sourceRange = $"A1:B{lastRow}";

            // Create a PivotTable that will serve as the data source for the Timeline
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIndex = pivots.Add(sourceRange, "D1", "EventPivot");
            PivotTable pivot = pivots[pivotIndex];

            // Use the "Date" column as the row field (time axis)
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            // Use the "Event" column as the data field (count of events)
            pivot.AddFieldToArea(PivotFieldType.Data, "Event");

            // Refresh and calculate the PivotTable data
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline linked to the PivotTable, positioned at cell A20 (row 19, column 0)
            int timelineIndex = sheet.Timelines.Add(pivot, 19, 0, "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];
            timeline.Caption = "Event Timeline";

            // Save the workbook with the Timeline control
            workbook.Save("EventTimeline.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
