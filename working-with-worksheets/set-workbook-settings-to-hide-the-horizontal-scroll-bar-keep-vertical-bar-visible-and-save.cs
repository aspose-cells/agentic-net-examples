// Title: Hide Horizontal Scroll Bar, Keep Vertical Scroll Bar Visible with Aspose.Cells for .NET
// Description: Shows how to create a workbook, set WorkbookSettings.IsHScrollBarVisible to false and IsVScrollBarVisible to true, and save the result as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | hide horizontal scrollbar | vertical scrollbar visible | WorkbookSettings | Excel scroll bar visibility | SaveFormat.Xlsx | programmatic Excel UI | Excel generation
// Common Searches: Aspose.Cells hide horizontal scrollbar C# | set workbook scroll bar visibility .NET | show only vertical scroll bar in generated Excel | Aspose.Cells workbook settings example | C# hide horizontal scroll bar in Excel file
// Developer Intent: Create an Excel workbook, disable the horizontal scroll bar while keeping the vertical scroll bar active, and save the file programmatically.
// Use Cases: Embedding generated Excel files in web portals where horizontal scrolling would break layout. | Designing mobile‑friendly reports that require only vertical navigation. | Preparing printable templates that should not allow sideways scrolling.
// AI Prompts: Generate C# code with Aspose.Cells that hides the horizontal scroll bar and customizes its width. | Explain how to toggle scroll bar visibility at runtime based on user settings using Aspose.Cells. | Show how to apply the same scroll bar configuration to an existing workbook loaded from disk.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, set WorkbookSettings.IsHScrollBarVisible to false and IsVScrollBarVisible to true, and save the result as an XLSX file using Aspose.Cells for .NET.
    public class WorkbookSettingsScrollBarDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access workbook settings
                WorkbookSettings settings = workbook.Settings;

                // Hide the horizontal scroll bar
                settings.IsHScrollBarVisible = false;

                // Keep the vertical scroll bar visible
                settings.IsVScrollBarVisible = true;

                // Save the workbook
                workbook.Save("ScrollBarSettings.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public static class Program
    {
        public static void Main()
        {
            WorkbookSettingsScrollBarDemo.Run();
        }
    }
}
