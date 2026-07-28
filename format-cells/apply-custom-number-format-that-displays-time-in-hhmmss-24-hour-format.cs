// Title: Aspose.Cells for .NET – Set hh:mm:ss 24‑hour time format on a cell (C#)
// Description: This C# example creates a workbook, writes a DateTime value (14:30:45) to cell A1, defines a style whose number format is "hh:mm:ss", applies the style, and saves the file as CustomTimeFormatDemo.xlsx using Aspose.Cells.
// Keywords: Aspose.Cells | C# time format | hh:mm:ss Excel | custom number format .NET | format cell as time | Excel 24 hour display | Aspose.Cells example | GitHub Aspose.Cells C#
// Common Searches: C# Aspose.Cells set cell time format | how to display hh:mm:ss in Excel with Aspose | apply custom number format to Excel cell .NET | Aspose.Cells time only formatting | save workbook with time format C#
// Developer Intent: I need to present a DateTime value in an Excel cell using a 24‑hour clock (hh:mm:ss) while keeping the underlying value unchanged.
// Use Cases: Generating production shift logs where each entry must show hour‑minute‑second. | Exporting telemetry timestamps for analytics dashboards that require a fixed 24‑hour display. | Creating timetable sheets for transportation schedules that need consistent time formatting across locales. | Automating report files where downstream systems parse time strings in hh:mm:ss format.
// AI Prompts: Write C# code that formats an entire column with hh:mm:ss using Aspose.Cells and keeps the style after the workbook is opened in Excel. | Demonstrate how to combine a custom time pattern with culture‑aware date parsing in Aspose.Cells for .NET. | Show how to apply the same 24‑hour format to a named range and then protect the worksheet. | Explain how to verify that the time format persists when the file is opened on different operating systems.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, writes a DateTime value (14:30:45) to cell A1, defines a style whose number format is "hh:mm:ss", applies the style, and saves the file as CustomTimeFormatDemo.xlsx using Aspose.Cells.
    public class CustomTimeFormatDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Insert a time value (14:30:45) into cell A1
            cells["A1"].PutValue(new DateTime(2023, 1, 1, 14, 30, 45));

            // Create a style and set a custom 24‑hour time format
            Style style = workbook.CreateStyle();
            style.Custom = "hh:mm:ss";

            // Apply the style to the cell
            cells["A1"].SetStyle(style);

            // Save the workbook
            workbook.Save("CustomTimeFormatDemo.xlsx");
        }
    }
}
