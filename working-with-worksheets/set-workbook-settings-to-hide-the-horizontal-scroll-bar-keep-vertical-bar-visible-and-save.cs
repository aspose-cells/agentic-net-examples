// Title: C# – Hide Horizontal Scroll Bar, Keep Vertical Scroll Bar Visible in Aspose.Cells Workbook and Save as XLSX
// Description: Demonstrates how to use Aspose.Cells WorkbookSettings in C# to disable the horizontal scroll bar, keep the vertical scroll bar enabled, and save the workbook in XLSX format.
// Keywords: Aspose.Cells C# hide horizontal scroll bar | WorkbookSettings scroll bar visibility | IsHScrollBarVisible false | IsVScrollBarVisible true | disable horizontal scroll bar Aspose.Cells | customize Excel UI with Aspose.Cells | save workbook as XLSX C# | Excel scroll bar settings programmatically
// Common Searches: how to hide horizontal scroll bar in Aspose.Cells C# | Aspose.Cells set scroll bar visibility | C# hide horizontal scroll bar Excel workbook | keep vertical scroll bar visible Aspose.Cells | save workbook with custom scroll bar settings
// Developer Intent: Disable the horizontal scroll bar, retain the vertical scroll bar, and save the workbook using Aspose.Cells in C#.
// Use Cases: Create clean‑look reports where only vertical navigation is needed. | Prepare workbooks for web viewers that require a fixed horizontal view. | Standardize exported Excel files by programmatically setting scroll bar visibility before saving.
// AI Prompts: Show C# code to hide the horizontal scroll bar while keeping the vertical scroll bar visible in an Aspose.Cells workbook. | Explain how to toggle scroll bar visibility settings with Aspose.Cells WorkbookSettings for multiple workbooks. | Provide a step‑by‑step guide to configure scroll bar visibility and save the file as XLSX using Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to use Aspose.Cells WorkbookSettings in C# to disable the horizontal scroll bar, keep the vertical scroll bar enabled, and save the workbook in XLSX format.
public class WorkbookScrollBarSettingsDemo
{
    public static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access workbook settings
        WorkbookSettings settings = workbook.Settings;

        // Hide the horizontal scroll bar
        settings.IsHScrollBarVisible = false;

        // Keep the vertical scroll bar visible
        settings.IsVScrollBarVisible = true;

        // Save the workbook in XLSX format
        workbook.Save("WorkbookWithScrollBars.xlsx", SaveFormat.Xlsx);
    }
}
