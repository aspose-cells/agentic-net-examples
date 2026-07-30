// Title: Display Horizontal Scroll Bar, Hide Vertical Scroll Bar, and Save Workbook as XLSX with Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, sets WorkbookSettings.IsHScrollBarVisible to true and IsVScrollBarVisible to false, then saves the file as ScrollBarDemo.xlsx (XLSX) using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# scroll bar visibility | WorkbookSettings horizontal scroll bar | hide vertical scroll bar Aspose.Cells | save workbook as XLSX | Aspose.Cells .NET example | Excel scroll bar settings C# | Aspose.Cells workbook settings
// Common Searches: Aspose.Cells how to show only horizontal scroll bar | C# hide vertical scroll bar in Excel workbook | set scroll bar visibility Aspose.Cells | save workbook to XLSX after changing settings | WorkbookSettings scrollbars Aspose.Cells example
// Developer Intent: Make the horizontal scroll bar visible, hide the vertical scroll bar, and write the workbook to an XLSX file.
// Use Cases: Design a web‑based spreadsheet viewer that restricts vertical scrolling while allowing horizontal navigation of wide tables. | Generate a printable report where vertical movement is disabled to preserve layout before exporting to XLSX. | Create a data‑heavy dashboard that scrolls only horizontally, then distribute the workbook to end users.
// AI Prompts: Write C# code using Aspose.Cells to enable the horizontal scroll bar, disable the vertical scroll bar, and save the workbook as XLSX. | Explain how to apply scroll bar visibility settings to all worksheets in an existing workbook with Aspose.Cells. | Show how to toggle scroll bar visibility based on user input and export the workbook in XLSX, PDF, and CSV formats.

using System;
using Aspose.Cells;

// Creates a new Workbook, sets WorkbookSettings.IsHScrollBarVisible to true and IsVScrollBarVisible to false, then saves the file as ScrollBarDemo.xlsx (XLSX) using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the workbook settings
        WorkbookSettings settings = workbook.Settings;

        // Ensure the horizontal scroll bar is visible
        settings.IsHScrollBarVisible = true;

        // Hide the vertical scroll bar
        settings.IsVScrollBarVisible = false;

        // Save the workbook to disk in XLSX format
        workbook.Save("ScrollBarDemo.xlsx", SaveFormat.Xlsx);
    }
}
