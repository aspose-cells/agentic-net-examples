// Title: Log previous and new worksheet page‑setup dimensions when changing to A4 with Aspose.Cells for .NET
// Common Searches: how to record worksheet page setup dimensions before and after setting PaperSize to A4 using Aspose.Cells in C# | append page setup change log for each sheet in an Excel workbook with Aspose.Cells .NET | capture previous PaperWidth and PaperHeight values when modifying page setup in Aspose.Cells | log worksheet page setup changes to a file while converting workbook to A4 size in C# | track paper size modifications per worksheet using Aspose.Cells for .NET
// Tags: Aspose.Cells worksheet PaperSize A4 logging | C# log previous and new paper dimensions | Aspose.Cells page setup dimension tracking | append Excel sheet change entries to log | record worksheet paper width height Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook (or creates a new one), iterates through each worksheet, records the current PaperWidth and PaperHeight, sets the page setup to A4, records the new dimensions, and writes both the previous and new measurements with timestamps to a log file before saving the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load workbook – create a new one if the input file does not exist
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook
                Console.WriteLine($"Input file not found. A new workbook has been created.");
            }

            // Prepare the log file (append mode)
            string logPath = "PageSetupChanges.log";
            using (StreamWriter logWriter = new StreamWriter(logPath, true))
            {
                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    try
                    {
                        // Capture previous paper dimensions (in millimeters)
                        double previousWidth = sheet.PageSetup.PaperWidth;
                        double previousHeight = sheet.PageSetup.PaperHeight;

                        // Change page setup to A4 size
                        sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

                        // Capture new paper dimensions
                        double newWidth = sheet.PageSetup.PaperWidth;
                        double newHeight = sheet.PageSetup.PaperHeight;

                        // Log the change details
                        logWriter.WriteLine($"{DateTime.Now}: Worksheet '{sheet.Name}' page‑setup changed.");
                        logWriter.WriteLine($"    Previous dimensions: Width={previousWidth}mm, Height={previousHeight}mm");
                        logWriter.WriteLine($"    New dimensions:      Width={newWidth}mm, Height={newHeight}mm");
                    }
                    catch (Exception ex)
                    {
                        logWriter.WriteLine($"{DateTime.Now}: Error processing worksheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
