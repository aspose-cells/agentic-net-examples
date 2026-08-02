// Title: C# – Show vertical scroll bar, hide horizontal scroll bar in Aspose.Cells workbook and save as XLSX
// Description: Creates a new Workbook, uses WorkbookSettings to enable the vertical scroll bar (IsVScrollBarVisible = true) and disable the horizontal scroll bar (IsHScrollBarVisible = false), then saves the file in XLSX format.
// Keywords: Aspose.Cells C# scroll bar visibility | show vertical scroll bar Aspose.Cells | hide horizontal scroll bar Aspose.Cells | WorkbookSettings IsVScrollBarVisible | WorkbookSettings IsHScrollBarVisible | save workbook as XLSX Aspose.Cells | Excel scroll bar control C# | Aspose.Cells example scroll bars
// Common Searches: Aspose.Cells hide horizontal scroll bar C# | Enable only vertical scroll bar in Excel file using Aspose.Cells | WorkbookSettings scroll bar visibility Aspose.Cells example | C# code to set scroll bar visibility and save workbook | How to hide horizontal scroll bar with Aspose.Cells
// Developer Intent: Configure workbook scroll bars (vertical visible, horizontal hidden) and export the workbook.
// Use Cases: Generate a read‑only report where users scroll vertically but cannot scroll horizontally. | Prepare a web‑viewer template that requires only vertical navigation before publishing the XLSX file. | Create a printable Excel sheet where horizontal scrolling is unnecessary, improving user focus.
// AI Prompts: Write C# code with Aspose.Cells to show the vertical scroll bar, hide the horizontal scroll bar, and save the workbook as XLSX. | Provide an Aspose.Cells example that modifies WorkbookSettings to control scroll bar visibility and then exports the file. | Explain the impact of changing IsVScrollBarVisible and IsHScrollBarVisible on different Excel viewers.

using System;
using Aspose.Cells;

namespace AsposeCellsScrollBarDemo
{
    // Creates a new Workbook, uses WorkbookSettings to enable the vertical scroll bar (IsVScrollBarVisible = true) and disable the horizontal scroll bar (IsHScrollBarVisible = false), then saves the file in XLSX format.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the workbook settings
            WorkbookSettings settings = workbook.Settings;

            // Show the vertical scroll bar
            settings.IsVScrollBarVisible = true;

            // Hide the horizontal scroll bar
            settings.IsHScrollBarVisible = false;

            // Save the workbook to disk (XLSX format)
            workbook.Save("ScrollBarDemo.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved with vertical scroll bar visible and horizontal scroll bar hidden.");
        }
    }
}
