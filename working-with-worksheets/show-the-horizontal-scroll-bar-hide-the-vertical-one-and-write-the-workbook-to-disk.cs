// Title: Show Horizontal Scroll Bar, Hide Vertical Scroll Bar, and Save Workbook – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to enable the horizontal scroll bar, disable the vertical scroll bar, add sample data, and save a workbook as XLSX using Aspose.Cells for .NET (C#).
// Keywords: Aspose.Cells C# | WorkbookSettings | IsHScrollBarVisible | IsVScrollBarVisible | hide vertical scroll bar | show horizontal scroll bar | save workbook as XLSX | C# Excel scroll bar visibility
// Common Searches: Aspose.Cells show only horizontal scroll bar C# | Hide vertical scroll bar in Aspose.Cells workbook | C# set scroll bar visibility Aspose.Cells | Save workbook after changing scroll bar settings Aspose.Cells | WorkbookSettings scroll bar C# example
// Developer Intent: Enable the horizontal scroll bar, hide the vertical scroll bar, and write the workbook to an XLSX file using Aspose.Cells for .NET.
// Use Cases: Create Excel dashboards where vertical scrolling is locked for a fixed layout. | Generate printable reports that require only horizontal navigation. | Prepare template files with predefined UI settings for end‑user distribution. | Build web‑based spreadsheet viewers that limit vertical scroll for better UX.
// AI Prompts: Write C# code with Aspose.Cells to make the horizontal scroll bar visible, hide the vertical scroll bar, and save the workbook as an XLSX file. | Explain the effect of WorkbookSettings.IsHScrollBarVisible and IsVScrollBarVisible on the Excel UI and how they persist after saving. | Suggest additional formatting (e.g., column width, freeze panes) that complements scroll bar settings in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsScrollBarDemo
{
    // Demonstrates how to enable the horizontal scroll bar, disable the vertical scroll bar, add sample data, and save a workbook as XLSX using Aspose.Cells for .NET (C#).
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access workbook settings
            WorkbookSettings settings = workbook.Settings;

            // Ensure the horizontal scroll bar is visible
            settings.IsHScrollBarVisible = true;

            // Hide the vertical scroll bar
            settings.IsVScrollBarVisible = false;

            // (Optional) Add some sample data to demonstrate the workbook
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["B1"].PutValue("Data");

            // Save the workbook to disk in XLSX format
            workbook.Save("ScrollBarDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
