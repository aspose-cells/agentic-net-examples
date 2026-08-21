// Title: Log PageSetup Paper Width & Height Changes with Aspose.Cells for .NET
// Description: Shows how to capture the initial, A4, Letter, and custom (2.5" × 3.0") paper dimensions of a worksheet, write timestamps and sizes to a log file, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | PageSetup | PaperWidth | PaperHeight | logging | custom paper size | PaperSizeType | workbook save | .NET | log file | dimension audit
// Common Searches: Aspose.Cells log page setup dimensions | record paper size changes in C# Aspose.Cells | how to write PageSetup PaperWidth to file | custom paper size logging Aspose.Cells .NET | save workbook after page setup changes
// Developer Intent: Record each PageSetup modification and its paper dimensions in a log file.
// Use Cases: Capture the worksheet's original paper width and height before any changes. | Log dimensions after switching to predefined sizes such as A4 and Letter. | Document the size resulting from a custom paper size (e.g., 2.5" × 3.0"). | Create an audit trail for page‑setup adjustments in automated report generation.
// AI Prompts: Generate C# code that logs PageSetup.PaperWidth and PaperHeight before and after every change using Aspose.Cells. | Extend the example to output the log in CSV format with columns for timestamp, action, width, and height. | Explain how to intercept all PageSetup property changes automatically without calling a helper method for each.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPageSetupLogging
{
    // Shows how to capture the initial, A4, Letter, and custom (2.5" × 3.0") paper dimensions of a worksheet, write timestamps and sizes to a log file, and save the workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (using the create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Access the PageSetup object
            PageSetup pageSetup = sheet.PageSetup;

            // Prepare a log file
            string logPath = "PageSetupChanges.log";
            using (StreamWriter log = new StreamWriter(logPath, false))
            {
                // Log initial paper dimensions
                LogDimensions(log, "Initial", pageSetup);

                // Change paper size to A4 using PaperSize property
                pageSetup.PaperSize = PaperSizeType.PaperA4;
                LogDimensions(log, "After setting PaperSize to A4", pageSetup);

                // Change paper size to Letter using PaperSize property
                pageSetup.PaperSize = PaperSizeType.PaperLetter;
                LogDimensions(log, "After setting PaperSize to Letter", pageSetup);

                // Set a custom paper size (width: 2.5 inches, height: 3.0 inches)
                pageSetup.CustomPaperSize(2.5, 3.0);
                LogDimensions(log, "After setting CustomPaperSize (2.5\" x 3.0\")", pageSetup);
            }

            // Save the workbook (using the save rule)
            workbook.Save("PageSetupLoggingDemo.xlsx");
        }

        // Helper method to write previous and new dimensions to the log
        static void LogDimensions(StreamWriter log, string action, PageSetup setup)
        {
            // PaperWidth and PaperHeight are read‑only, they reflect the current size
            double width = setup.PaperWidth;   // in inches
            double height = setup.PaperHeight; // in inches

            log.WriteLine($"{DateTime.Now:u} - {action}");
            log.WriteLine($"    Paper Width : {width:F2} inches");
            log.WriteLine($"    Paper Height: {height:F2} inches");
            log.WriteLine();
        }
    }
}
