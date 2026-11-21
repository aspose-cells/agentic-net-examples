// Title: Log each worksheet's page‑setup changes, including previous and new paper dimensions, to a file using Aspose.Cells for .NET
// Common Searches: how to write a log of page setup changes for every worksheet with Aspose.Cells C# | record previous paper width and height before setting A4 size in Excel using Aspose.Cells | Aspose.Cells C# example to audit worksheet page setup and save to a text file | save worksheet page setup dimensions to a log file when changing paper size with Aspose.Cells | C# code to capture and log old and new paper size values for Excel worksheets using Aspose.Cells
// Tags: Aspose.Cells worksheet page setup audit | C# log Excel paper dimensions | Aspose.Cells change paper size to A4 | record previous and new page size values | write page setup details to text file

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an existing or new workbook, creates necessary log and output folders, iterates through each worksheet, captures the current PaperWidth, PaperHeight, and PaperSize, sets the page setup to A4, writes the before‑and‑after dimensions to a log file, and finally saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string logFilePath = "PageSetupChanges.log";

        Workbook workbook = null;

        // Load existing workbook or create a new one
        try
        {
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // default workbook with one worksheet
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Ensure log directory exists
        try
        {
            string logDir = Path.GetDirectoryName(logFilePath);
            if (!string.IsNullOrEmpty(logDir) && !Directory.Exists(logDir))
                Directory.CreateDirectory(logDir);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error preparing log directory: {ex.Message}");
            return;
        }

        // Process worksheets and write log
        try
        {
            using (StreamWriter logWriter = new StreamWriter(logFilePath, false))
            {
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    PageSetup pageSetup = sheet.PageSetup;

                    double previousWidth = pageSetup.PaperWidth;      // inches
                    double previousHeight = pageSetup.PaperHeight;    // inches
                    PaperSizeType previousPaperSize = pageSetup.PaperSize;

                    // Set A4 paper size (standard dimensions)
                    pageSetup.PaperSize = PaperSizeType.PaperA4;
                    // Aspose.Cells automatically adjusts width/height for standard sizes

                    // Log the changes
                    logWriter.WriteLine($"Worksheet: {sheet.Name}");
                    logWriter.WriteLine($"Previous PaperSize: {previousPaperSize}  -> New PaperSize: {pageSetup.PaperSize}");
                    logWriter.WriteLine($"Previous Width (in): {previousWidth:F2} -> New Width (in): {pageSetup.PaperWidth:F2}");
                    logWriter.WriteLine($"Previous Height (in): {previousHeight:F2} -> New Height (in): {pageSetup.PaperHeight:F2}");
                    logWriter.WriteLine(new string('-', 60));
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing worksheets or writing log: {ex.Message}");
            return;
        }

        // Ensure output directory exists before saving
        try
        {
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error preparing output directory: {ex.Message}");
            return;
        }

        // Save the modified workbook
        try
        {
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
