// Title: C# – Convert CSV to Excel and Add an Interactive Timeline with Aspose.Cells
// Description: A complete C# example that checks for a CSV file of events, creates a sample if missing, converts the CSV to XLSX using Aspose.Cells ConversionUtility, builds a pivot table on the Event and Date columns, inserts a linked timeline, and saves the workbook with the timeline for easy date navigation.
// Keywords: Aspose.Cells CSV to XLSX | C# timeline worksheet | Aspose.Cells pivot table | add timeline to Excel | ConversionUtility example | .NET Excel automation | interactive timeline Aspose | event timeline Excel | CSV import Aspose.Cells
// Common Searches: how to convert csv to excel with Aspose.Cells C# | add a timeline to a worksheet using Aspose.Cells | create pivot table and timeline from CSV in .NET | Aspose.Cells example for event timeline | C# code to generate Excel timeline from CSV
// Developer Intent: Import event dates from a CSV file, transform them into an Excel workbook, and visualize the dates with an interactive timeline linked to a pivot table.
// Use Cases: Generate a project‑milestone workbook from a CSV source with a clickable timeline for quick date filtering. | Produce a sales‑events report that converts raw CSV data to Excel, creates a pivot table, and adds a timeline for chronological analysis. | Automate an event‑log workbook that includes a pivot table and an interactive timeline to visualize and explore chronological data.
// AI Prompts: Provide a C# snippet that reads a CSV, converts it to XLSX with Aspose.Cells ConversionUtility, creates a pivot table, and adds a timeline linked to the Date field. | Explain how to customize the timeline’s style, size, and position after inserting it into a worksheet using Aspose.Cells. | Suggest best practices for handling large CSV files when building a pivot table and timeline with Aspose.Cells, including memory‑management tips.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Utility;

// A complete C# example that checks for a CSV file of events, creates a sample if missing, converts the CSV to XLSX using Aspose.Cells ConversionUtility, builds a pivot table on the Event and Date columns, inserts a linked timeline, and saves the workbook with the timeline for easy date navigation.
class TimelineFromCsv
{
    static void Main()
    {
        // Paths for source CSV and intermediate/final Excel files
        string csvPath = "events.csv";
        string intermediateXlsx = "events.xlsx";
        string finalXlsx = "events_with_timeline.xlsx";

        try
        {
            // Ensure the CSV file exists; create a simple sample if it does not.
            if (!File.Exists(csvPath))
            {
                // Sample data: Event,Date
                string[] sampleLines =
                {
                    "Event,Date",
                    "Launch,2023-01-15",
                    "Update,2023-02-10",
                    "Release,2023-03-05"
                };
                File.WriteAllLines(csvPath, sampleLines);
                Console.WriteLine($"Sample CSV created at '{Path.GetFullPath(csvPath)}'.");
            }

            // 1. Convert CSV to XLSX using Aspose.Cells ConversionUtility
            ConversionUtility.Convert(csvPath, intermediateXlsx);

            // Verify the intermediate file was created
            if (!File.Exists(intermediateXlsx))
                throw new FileNotFoundException($"Failed to create intermediate file '{intermediateXlsx}'.");

            // 2. Load the newly created workbook
            Workbook workbook = new Workbook(intermediateXlsx);
            Worksheet sheet = workbook.Worksheets[0];

            // 3. Create a PivotTable that uses the imported data.
            //    Assuming the CSV has columns "Event" (A) and "Date" (B).
            //    Determine the used range to set an accurate data area.
            int lastRow = sheet.Cells.MaxDataRow + 1; // +1 because rows are zero‑based
            string dataRange = $"A1:B{lastRow}";
            int pivotIndex = sheet.PivotTables.Add(dataRange, "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add the Date field as a row (timeline base) and Event as data (count)
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Event");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

            // Refresh and calculate the pivot data
            pivot.RefreshData();
            pivot.CalculateData();

            // 4. Add a Timeline linked to the PivotTable, using the "Date" field.
            //    The timeline will be placed starting at cell A20 (row 19, column 0).
            int timelineIndex = sheet.Timelines.Add(pivot, 19, 0, "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];
            timeline.Caption = "Event Timeline";

            // 5. Save the final workbook with the timeline
            workbook.Save(finalXlsx, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully as '{finalXlsx}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
