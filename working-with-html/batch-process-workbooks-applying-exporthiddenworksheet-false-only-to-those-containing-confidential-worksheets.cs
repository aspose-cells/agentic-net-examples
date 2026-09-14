// Title: Batch hide confidential worksheets in multiple Excel workbooks using Aspose.Cells for .NET
// AI Prompts: Generate a C# console application that scans a directory for .xlsx files, loads each workbook with Aspose.Cells, hides any worksheet whose name contains the word “confidential” (case‑insensitive), and saves the modified files to a separate output folder. | Write a .NET script that processes all Excel files in a given folder, sets Worksheet.IsVisible = false for sheets matching a confidential keyword, handles missing files and exceptions, and logs the processing results.
// Common Searches: C# Aspose.Cells hide worksheets containing a specific keyword in batch | How to process multiple Excel files and hide confidential sheets with Aspose.Cells | Automate hiding of confidential worksheets across many workbooks using .NET | Aspose.Cells hide sheet by name while saving multiple workbooks
// Tags: batch hide confidential worksheets Aspose.Cells | process multiple .xlsx workbooks C# | set worksheet IsVisible false Aspose.Cells | hide sheets by name keyword .NET | automate Excel workbook privacy Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example enumerates all .xlsx files in a source folder, loads each workbook with Aspose.Cells, marks any worksheet whose name includes the word "confidential" (case‑insensitive) as invisible, and saves the updated workbook to a target directory, creating the directory if needed and logging any errors encountered.
class BatchWorkbookProcessor
{
    static void Main(string[] args)
    {
        // Input and output directories
        string inputFolder = @"C:\InputWorkbooks";
        string outputFolder = @"C:\OutputWorkbooks";

        // Ensure output directory exists
        Directory.CreateDirectory(outputFolder);

        // Get all Excel files in the input folder
        string[] files = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

        foreach (string inputPath in files)
        {
            try
            {
                // Verify the file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"File not found: {inputPath}");
                    continue;
                }

                // Load workbook
                Workbook workbook = new Workbook(inputPath);

                // Hide confidential worksheets (if any)
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Name.IndexOf("confidential", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        sheet.IsVisible = false; // hide the sheet
                    }
                }

                // Build output file path
                string fileName = Path.GetFileName(inputPath);
                string outputPath = Path.Combine(outputFolder, fileName);

                // Save workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Processed and saved: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{inputPath}': {ex.Message}");
            }
        }
    }
}
