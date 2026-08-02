// Title: C# – Log Aspose.Cells PageSetup Changes with Old and New Paper Dimensions
// Description: A C# console app that creates a workbook, captures the initial PageSetup dimensions, then modifies the paper size (A4), applies a custom size (2.5" × 3.5"), switches orientation to Landscape, and writes a timestamped entry for each change—including previous and new width/height—to a log file. The workbook is saved and a confirmation message is displayed.
// Keywords: Aspose.Cells | C# | PageSetup logging | paper size tracking | custom paper size | orientation change | Excel worksheet audit | log file output | Workbook lifecycle | debug page layout
// Common Searches: how to log page setup changes in Aspose.Cells C# | track paper width and height after PageSetup modifications | Aspose.Cells write page layout changes to a text file | C# log worksheet orientation and paper size with Aspose.Cells | audit Excel page setup using Aspose.Cells
// Developer Intent: Record every PageSetup modification—paper size, custom dimensions, and orientation—by appending old and new measurements to a log file.
// Use Cases: Maintain a change history for printing settings when generating automated Excel reports. | Debug unexpected layout shifts during workbook creation by comparing logged dimensions. | Provide compliance evidence for document‑printing standards in regulated industries.
// AI Prompts: Create a reusable Aspose.Cells logger that captures any PageSetup property change with timestamps, old value, and new value. | Extend the logger to include margins, scaling, and header/footer settings for each worksheet in a workbook. | Generate code that routes PageSetup change logs to multiple destinations (file, database, or Windows Event Log) based on a configuration file.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPageSetupLogging
{
    // A C# console app that creates a workbook, captures the initial PageSetup dimensions, then modifies the paper size (A4), applies a custom size (2.5" × 3.5"), switches orientation to Landscape, and writes a timestamped entry for each change—including previous and new width/height—to a log file. The workbook is saved and a confirmation message is displayed.
    class Program
    {
        // Path to the log file
        private const string LogFilePath = "PageSetupChanges.log";

        static void Main()
        {
            // Ensure the log file is empty at start
            File.WriteAllText(LogFilePath, string.Empty);

            // Create a new workbook (using the lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Access the PageSetup object
            PageSetup pageSetup = sheet.PageSetup;

            // Log the initial paper dimensions
            LogPageSetupChange("Initial", pageSetup.PaperWidth, pageSetup.PaperHeight);

            // -----------------------------------------------------------------
            // Change 1: Set a predefined paper size (A4)
            // -----------------------------------------------------------------
            // Store previous dimensions
            double prevWidth = pageSetup.PaperWidth;
            double prevHeight = pageSetup.PaperHeight;

            // Apply new paper size
            pageSetup.PaperSize = PaperSizeType.PaperA4;

            // Log the change with previous and new dimensions
            LogPageSetupChange("PaperSize -> A4", prevWidth, prevHeight, pageSetup.PaperWidth, pageSetup.PaperHeight);

            // -----------------------------------------------------------------
            // Change 2: Set a custom paper size (2.5 x 3.5 inches)
            // -----------------------------------------------------------------
            prevWidth = pageSetup.PaperWidth;
            prevHeight = pageSetup.PaperHeight;

            // Apply custom size (width, height in inches)
            pageSetup.CustomPaperSize(2.5, 3.5);

            // Log the custom size change
            LogPageSetupChange("CustomPaperSize (2.5\" x 3.5\")", prevWidth, prevHeight, pageSetup.PaperWidth, pageSetup.PaperHeight);

            // -----------------------------------------------------------------
            // Change 3: Change orientation to Landscape (affects width/height)
            // -----------------------------------------------------------------
            prevWidth = pageSetup.PaperWidth;
            prevHeight = pageSetup.PaperHeight;

            pageSetup.Orientation = PageOrientationType.Landscape;

            // After changing orientation, PaperWidth and PaperHeight are swapped
            LogPageSetupChange("Orientation -> Landscape", prevWidth, prevHeight, pageSetup.PaperWidth, pageSetup.PaperHeight);

            // Save the workbook (using the lifecycle rule: save)
            workbook.Save("PageSetupLoggingDemo.xlsx");

            Console.WriteLine("Page setup changes have been logged to " + LogFilePath);
        }

        /// <param name="action">Description of the change.</param>
        /// <param name="oldWidth">Previous paper width (in inches).</param>
        /// <param name="oldHeight">Previous paper height (in inches).</param>
        /// <param name="newWidth">New paper width (in inches). If null, old values are used.</param>
        /// <param name="newHeight">New paper height (in inches). If null, old values are used.</param>
        private static void LogPageSetupChange(string action, double oldWidth, double oldHeight, double? newWidth = null, double? newHeight = null)
        {
            string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {action} | " +
                              $"Prev Width: {oldWidth:F3} in, Prev Height: {oldHeight:F3} in";

            if (newWidth.HasValue && newHeight.HasValue)
            {
                logEntry += $", New Width: {newWidth.Value:F3} in, New Height: {newHeight.Value:F3} in";
            }

            // Append the entry to the log file
            File.AppendAllText(LogFilePath, logEntry + Environment.NewLine);
        }
    }
}
