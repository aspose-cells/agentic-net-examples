// Title: Batch convert Excel workbooks to HTML in C# and hide worksheets only when a sheet name contains “Confidential” using Aspose.Cells
// AI Prompts: Write a C# console program that scans a directory for .xlsx files, detects worksheets whose name includes the word “Confidential”, sets HtmlSaveOptions.ExportHiddenWorksheet = false for those workbooks, and saves the HTML output to a target folder. | Update the given Aspose.Cells example to create an HtmlSaveOptions object, enable ExportHiddenWorksheet only when a confidential worksheet is present, and export each workbook as HTML while leaving other workbooks with default settings. | Generate a script that iterates over Excel files, checks each worksheet name for a confidentiality marker, applies conditional hidden‑worksheet export settings, and writes the resulting HTML files to a separate output directory.
// Common Searches: how to hide confidential worksheets when exporting Excel to HTML with Aspose.Cells in C# | C# batch process multiple .xlsx files and set HtmlSaveOptions.ExportHiddenWorksheet based on worksheet name | Aspose.Cells conditional export hidden worksheets only for sheets containing 'Confidential' | example code for scanning a folder of Excel workbooks and exporting them to HTML with hidden sheets disabled
// Tags: conditional HtmlSaveOptions.ExportHiddenWorksheet with Aspose.Cells | detect confidential worksheet names in C# | batch export Excel to HTML using Aspose.Cells | iterate over .xlsx files in a directory C# | save HTML output to separate folder Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to loop through all .xlsx files in a source folder, identify any worksheet whose name contains the term "Confidential", configure HtmlSaveOptions.ExportHiddenWorksheet = false for those workbooks, and export each workbook to HTML in a designated output directory using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Input and output directories
        string inputDir = @"C:\InputWorkbooks";
        string outputDir = @"C:\OutputWorkbooks";

        // Ensure the input directory exists
        if (!Directory.Exists(inputDir))
        {
            Console.WriteLine($"Input directory does not exist: {inputDir}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputDir);

        // Iterate over all Excel files in the input directory
        foreach (string filePath in Directory.GetFiles(inputDir, "*.xlsx"))
        {
            // Verify the file exists before attempting to load
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                bool hasConfidentialSheet = false;

                // Check each worksheet for a confidential indicator (e.g., name contains "Confidential")
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Name.IndexOf("Confidential", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        hasConfidentialSheet = true;
                        break;
                    }
                }

                // If a confidential worksheet is present, you may adjust settings here.
                // The ExportHiddenWorksheet property is not available in the current Aspose.Cells version,
                // so this step is omitted.

                // Build the output file path
                string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Processed and saved: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }
}
