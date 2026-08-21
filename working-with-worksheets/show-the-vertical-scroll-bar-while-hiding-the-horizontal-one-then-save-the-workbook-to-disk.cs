// Title: C# – Show vertical scroll bar, hide horizontal scroll bar in Aspose.Cells and save workbook
// Description: Creates a new Workbook, sets WorkbookSettings.IsVScrollBarVisible to true and IsHScrollBarVisible to false, ensures the output folder exists, and saves the file as an XLSX document.
// Keywords: Aspose.Cells vertical scroll bar | hide horizontal scroll bar Aspose.Cells | C# workbook settings scrollbars | save workbook with custom scroll bar visibility | Aspose.Cells example scroll bar control
// Common Searches: Aspose.Cells show only vertical scroll bar C# | hide horizontal scroll bar Aspose.Cells workbook | set scroll bar visibility Aspose.Cells | save Excel file after changing scroll bar settings | C# example for workbook scroll bar configuration
// Developer Intent: Enable the vertical scroll bar, disable the horizontal scroll bar, and persist the workbook to disk.
// Use Cases: Design a report where users scroll only vertically, preventing sideways navigation. | Prepare an Excel view for web embedding that requires a clean vertical‑only scroll experience. | Generate a print‑ready spreadsheet where horizontal scroll bars are unnecessary.
// AI Prompts: Generate C# code using Aspose.Cells to hide the horizontal scroll bar, show the vertical scroll bar, and save the workbook as XLSX. | Provide a reusable method that accepts a file path, creates a workbook with only the vertical scroll bar visible, and returns the saved file location. | Explain how to configure WorkbookSettings for scroll bar visibility and ensure the output directory exists before saving.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new Workbook, sets WorkbookSettings.IsVScrollBarVisible to true and IsHScrollBarVisible to false, ensures the output folder exists, and saves the file as an XLSX document.
    public class ShowVerticalHideHorizontalScrollBar
    {
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
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access workbook settings
            WorkbookSettings settings = workbook.Settings;

            // Show vertical scroll bar
            settings.IsVScrollBarVisible = true;

            // Hide horizontal scroll bar
            settings.IsHScrollBarVisible = false;

            // Ensure output directory exists
            string outputPath = "VerticalScrollBarOnly.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to disk (XLSX format)
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved with vertical scroll bar visible and horizontal scroll bar hidden.");
        }
    }
}
