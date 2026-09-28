// Title: Load events.csv into an Aspose.Cells workbook, build a PivotTable and attach a timeline control using C#
// AI Prompts: Read a CSV file with Aspose.Cells, import it to a worksheet, create a PivotTable on the imported range, and add a timeline linked to the date column in C#. | Generate an Excel workbook from events.csv, then configure a timeline slicer that syncs with the PivotTable's date field using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# import CSV and create timeline slicer from pivot table | how to add a timeline to a PivotTable in Aspose.Cells .NET | convert events.csv to Excel with timeline control using Aspose.Cells | C# example for CSV import, PivotTable creation, and timeline in Aspose.Cells | Aspose.Cells timeline feature example with CSV data
// Tags: import csv to worksheet Aspose.Cells | create pivot table from csv Aspose.Cells | add timeline slicer to pivot Aspose.Cells | save workbook with timeline xlsx Aspose.Cells | event date visualization Excel timeline .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

// The program reads events.csv, imports it into the first worksheet, builds a PivotTable on the data, adds a timeline slicer linked to the Date field, and saves the workbook as EventTimeline.xlsx.
class Program
{
    static void Main()
    {
        // Path to the CSV file that contains event dates.
        string csvPath = "events.csv";

        // Create a new workbook.
        Workbook workbook = new Workbook();

        // Access the first worksheet.
        Worksheet sheet = workbook.Worksheets[0];

        // Import the CSV data into the worksheet starting at cell A1.
        // Using comma as the delimiter and converting numeric data automatically.
        sheet.Cells.ImportCSV(csvPath, ",", true, 0, 0);

        // Add a PivotTable based on the imported data.
        // Adjust the range "A1:B100" to match the actual size of your CSV if needed.
        int pivotIndex = sheet.PivotTables.Add("A1:B100", "E1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Add fields to the PivotTable.
        // Assuming column 0 holds the event description and column 1 holds the date.
        pivot.AddFieldToArea(PivotFieldType.Row, 0);      // Event description
        pivot.AddFieldToArea(PivotFieldType.Column, 1);   // Date field
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a Timeline linked to the PivotTable.
        // The timeline will be placed with its upper‑left corner at cell G1
        // and will use the "Date" field as its base field.
        int timelineIndex = sheet.Timelines.Add(pivot, "G1", "Date");
        Timeline timeline = sheet.Timelines[timelineIndex];
        timeline.Caption = "Event Timeline";

        // Save the workbook containing the timeline.
        workbook.Save("EventTimeline.xlsx", SaveFormat.Xlsx);
    }
}
