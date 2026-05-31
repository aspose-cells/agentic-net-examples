using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPageSetupLogging
{
    class Program
    {
        static void Main()
        {
            // Path for the log file
            string logPath = "PageSetupChanges.log";

            // Ensure the log file is empty at start
            File.WriteAllText(logPath, string.Empty);

            // Create a new workbook (using the lifecycle rule for creation)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the PageSetup object
            PageSetup pageSetup = worksheet.PageSetup;

            // Helper method to log a change
            void LogChange(string changeDescription, double oldWidth, double oldHeight, double newWidth, double newHeight)
            {
                string logEntry = $"{DateTime.Now:u} | {changeDescription}\n" +
                                  $"    Previous Size (inches): Width={oldWidth:F2}, Height={oldHeight:F2}\n" +
                                  $"    New Size (inches):      Width={newWidth:F2}, Height={newHeight:F2}\n";
                File.AppendAllText(logPath, logEntry);
            }

            // -----------------------------------------------------------------
            // First change: set a predefined paper size (A4)
            // -----------------------------------------------------------------
            double prevWidth = pageSetup.PaperWidth;
            double prevHeight = pageSetup.PaperHeight;

            // Change paper size
            pageSetup.PaperSize = PaperSizeType.PaperA4;

            double newWidth = pageSetup.PaperWidth;
            double newHeight = pageSetup.PaperHeight;

            LogChange("Set PaperSize to A4", prevWidth, prevHeight, newWidth, newHeight);

            // -----------------------------------------------------------------
            // Second change: set a custom paper size (2.5 x 3.5 inches)
            // -----------------------------------------------------------------
            prevWidth = pageSetup.PaperWidth;
            prevHeight = pageSetup.PaperHeight;

            // Apply custom size
            pageSetup.CustomPaperSize(2.5, 3.5);

            newWidth = pageSetup.PaperWidth;
            newHeight = pageSetup.PaperHeight;

            LogChange("Set CustomPaperSize to 2.5\" x 3.5\"", prevWidth, prevHeight, newWidth, newHeight);

            // -----------------------------------------------------------------
            // Third change: change orientation to Landscape (affects width/height)
            // -----------------------------------------------------------------
            prevWidth = pageSetup.PaperWidth;
            prevHeight = pageSetup.PaperHeight;

            pageSetup.Orientation = PageOrientationType.Landscape;

            newWidth = pageSetup.PaperWidth;
            newHeight = pageSetup.PaperHeight;

            LogChange("Changed Orientation to Landscape", prevWidth, prevHeight, newWidth, newHeight);

            // Save the workbook (using the lifecycle rule for saving)
            workbook.Save("PageSetupLoggingDemo.xlsx");

            // Optional: output location of log file
            Console.WriteLine($"Page setup changes logged to: {Path.GetFullPath(logPath)}");
        }
    }
}