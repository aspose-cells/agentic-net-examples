// Title: Capture start and end timestamps for each worksheet enumeration pass using Aspose.Cells in C#
// AI Prompts: Insert DateTime.Now calls before and after the inner cell loop of each worksheet and add the values to a List<(string sheet, DateTime start, DateTime end)>. | Create a console output that formats the sheet name, start time, end time, and elapsed milliseconds for every entry in the performance list. | Modify the enumeration code to record timing information without altering any cell values or saving changes.
// Common Searches: how to measure worksheet enumeration time with Aspose.Cells in C# | record start and end DateTime for each sheet loop using Aspose.Cells .NET | performance logging of cell traversal per worksheet in Excel file via Aspose.Cells | benchmark sheet processing duration with Aspose.Cells API in C#
// Tags: sheet loop timing Aspose.Cells | cell iteration duration measurement .NET | worksheet processing time logging C# | excel workbook timestamp tracking Aspose.Cells | performance analysis of sheet enumeration

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example records DateTime.Now before and after iterating each worksheet's cells, stores the sheet name with start and end timestamps in a list, and prints the duration in milliseconds for performance analysis.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // List to store start and end timestamps for each enumeration pass
        List<(string SheetName, DateTime Start, DateTime End)> performanceLog = new List<(string, DateTime, DateTime)>();

        // Enumerate each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Record the start time of the enumeration pass
            DateTime startTime = DateTime.Now;

            // Example enumeration: iterate through all used cells in the worksheet
            foreach (Cell cell in sheet.Cells)
            {
                // Perform a lightweight operation (e.g., read the cell value)
                var value = cell.Value;
            }

            // Record the end time of the enumeration pass
            DateTime endTime = DateTime.Now;

            // Store the timestamps along with the sheet name
            performanceLog.Add((sheet.Name, startTime, endTime));
        }

        // Output the performance log
        foreach (var entry in performanceLog)
        {
            TimeSpan duration = entry.End - entry.Start;
            Console.WriteLine($"Sheet: {entry.SheetName}, Start: {entry.Start:O}, End: {entry.End:O}, Duration: {duration.TotalMilliseconds} ms");
        }

        // Save the workbook if any modifications were made
        workbook.Save("output.xlsx");
    }
}
