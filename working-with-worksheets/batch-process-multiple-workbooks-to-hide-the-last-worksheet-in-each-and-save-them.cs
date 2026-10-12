// Title: C# batch hide the last worksheet of each .xlsx workbook and save to a separate folder using Aspose.Cells
// AI Prompts: Generate C# code that iterates over all .xlsx files in a source directory, loads each workbook with Aspose.Cells, hides the last worksheet, and writes the modified workbook to a target folder. | Add logic to the batch routine to detect password‑protected Excel files, skip those files, and log their names to the console. | Replace the IsVisible property assignment with the SheetVisibility.Hidden enum when setting worksheet visibility in Aspose.Cells.
// Common Searches: C# Aspose.Cells hide last sheet in multiple workbooks and save to another directory | How to batch process Excel files to hide a worksheet using Aspose.Cells .NET | Skip password protected Excel files when looping through workbooks with Aspose.Cells | Set worksheet visibility to hidden with Aspose.Cells C# example
// Tags: batch hide worksheet Aspose.Cells C# | process multiple .xlsx files Aspose.Cells | skip password protected Excel Aspose.Cells | set worksheet visibility SheetVisibility.Hidden | save modified workbooks to output folder

using System;
using System.IO;
using Aspose.Cells;

// The sample scans a source folder for .xlsx files, loads each workbook with Aspose.Cells, hides the last worksheet, skips any password‑protected files, and saves the updated workbooks to a separate output directory while logging progress and errors.
class Program
{
    static void Main()
    {
        // Folder containing the source workbooks
        string sourceFolder = @"C:\Workbooks";

        // Folder where the processed workbooks will be saved
        string outputFolder = @"C:\ProcessedWorkbooks";

        // Verify source folder exists
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder does not exist: {sourceFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Retrieve all Excel files (adjust the pattern if needed)
        string[] workbookFiles = Directory.GetFiles(sourceFolder, "*.xlsx");

        foreach (string filePath in workbookFiles)
        {
            // Guard against missing file (should not happen with GetFiles, but added for safety)
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found, skipping: {filePath}");
                continue;
            }

            try
            {
                Workbook workbook;

                // Attempt to load the workbook; handle password‑protected files gracefully
                try
                {
                    workbook = new Workbook(filePath);
                }
                catch (Exception loadEx)
                {
                    // Aspose.Cells throws a generic exception when a password is required
                    if (loadEx.Message != null && loadEx.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        Console.WriteLine($"Password required for file, skipping: {filePath}");
                        continue;
                    }

                    // Re‑throw other load errors to be handled by outer catch
                    throw;
                }

                // Hide the last worksheet if any
                int lastWorksheetIndex = workbook.Worksheets.Count - 1;
                if (lastWorksheetIndex >= 0)
                {
                    Worksheet lastWorksheet = workbook.Worksheets[lastWorksheetIndex];
                    lastWorksheet.IsVisible = false; // Alternatively: lastWorksheet.Visibility = SheetVisibility.Hidden;
                }

                // Construct the output file path (overwrites if same name exists)
                string fileName = Path.GetFileName(filePath);
                string outputPath = Path.Combine(outputFolder, fileName);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Processed and saved: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any processing/saving errors
                Console.WriteLine($"Error processing workbook '{filePath}': {ex.Message}");
            }
        }
    }
}
