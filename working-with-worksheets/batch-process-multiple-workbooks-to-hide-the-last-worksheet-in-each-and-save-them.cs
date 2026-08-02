// Title: Batch hide the last worksheet in multiple Excel workbooks with Aspose.Cells for .NET
// Description: Iterates over a collection of Excel file paths, loads each workbook with Aspose.Cells, hides the final worksheet using SetVisible(false, true), and saves the updated file to a designated output folder while handling missing files and runtime exceptions.
// Keywords: Aspose.Cells hide worksheet C# | batch process Excel workbooks .NET | SetVisible false Aspose | hide last sheet programmatically | iterate multiple workbooks Aspose.Cells | save modified workbook Aspose | Excel automation hide sheet C# | process Excel files Aspose.Cells
// Common Searches: How to hide the last sheet in several Excel files using Aspose.Cells | Batch hide worksheets in .NET with Aspose.Cells | Hide a worksheet without error Aspose.Cells C# | Automate hiding sheets across multiple workbooks | Aspose.Cells hide sheet and save workbook
// Developer Intent: Loop through a set of Excel workbooks, hide each workbook’s last worksheet, and write the modified files to a target directory.
// Use Cases: Prepare distribution‑ready reports by hiding internal summary sheets before publishing. | Remove confidential worksheets from dozens of files during a data‑migration project. | Create publishable workbooks where the final sheet contains notes that must stay hidden from end users.
// AI Prompts: Generate a C# method that accepts a list of Excel file paths and hides the last worksheet in each using Aspose.Cells, preserving original filenames. | Add logging to the batch routine that records successful file names and error details to a CSV report. | Modify the method to include a Boolean flag that either hides or unhides the last worksheet based on the caller’s request.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace BatchWorkbookProcessing
{
    // Iterates over a collection of Excel file paths, loads each workbook with Aspose.Cells, hides the final worksheet using SetVisible(false, true), and saves the updated file to a designated output folder while handling missing files and runtime exceptions.
    public class HideLastWorksheetBatch
    {
        /// <param name="inputFiles">Full paths of the workbooks to process.</param>
        /// <param name="outputFolder">Folder where the modified workbooks will be saved.</param>
        public static void ProcessWorkbooks(IEnumerable<string> inputFiles, string outputFolder)
        {
            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            foreach (string inputPath in inputFiles)
            {
                try
                {
                    // Verify the source file exists before attempting to load
                    if (!File.Exists(inputPath))
                    {
                        Console.WriteLine($"File not found: {inputPath}. Skipping.");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(inputPath);

                    // Hide the last worksheet, if any
                    int lastIndex = workbook.Worksheets.Count - 1;
                    if (lastIndex >= 0)
                    {
                        // SetVisible(false, true) hides the sheet and ignores errors
                        workbook.Worksheets[lastIndex].SetVisible(false, true);
                    }

                    // Build the output file path (preserve original file name)
                    string outputPath = Path.Combine(outputFolder, Path.GetFileName(inputPath));

                    // Save the modified workbook
                    workbook.Save(outputPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{inputPath}': {ex.Message}");
                }
            }
        }

        // Example usage
        public static void Main()
        {
            // Define the list of workbook files to process
            List<string> files = new List<string>
            {
                @"C:\Data\Workbook1.xlsx",
                @"C:\Data\Workbook2.xlsx",
                @"C:\Data\Workbook3.xlsx"
            };

            // Specify the folder where the processed workbooks will be saved
            string outputDir = @"C:\Data\Processed";

            try
            {
                // Execute the batch operation
                ProcessWorkbooks(files, outputDir);
                Console.WriteLine("Batch processing completed. Modified files are saved in: " + outputDir);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unexpected error during batch processing: " + ex.Message);
            }
        }
    }
}
