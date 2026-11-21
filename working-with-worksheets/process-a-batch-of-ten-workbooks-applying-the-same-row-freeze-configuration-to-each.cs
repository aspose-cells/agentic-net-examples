// Title: Freeze the top row in up to ten Excel workbooks from a folder using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells batch freeze first row in multiple workbooks | How to apply FreezePanes to all worksheets in a directory of Excel files with Aspose.Cells | Process up to 10 Excel files and set row freeze while preserving original filenames in .NET | Load workbooks from a folder and freeze the top row programmatically using Aspose.Cells | Save modified workbooks to a separate output folder after applying freeze panes in C#
// Tags: freeze panes first row Aspose.Cells | batch workbook processing C# | apply FreezePanes to all worksheets | preserve original filenames Aspose.Cells | load and save workbooks from directory

using System;
using System.IO;
using Aspose.Cells;

namespace BatchFreezeRowsApp
{
    // The example scans an input directory, loads up to ten Excel workbooks of any supported format, iterates through each worksheet, freezes the first row (no column freeze) using Worksheet.FreezePanes, and saves the updated files to an output folder while keeping the original file names.
    class BatchFreezeRows
    {
        static void Main()
        {
            // Define input and output folders
            string inputFolder = @"C:\Workbooks\Input";
            string outputFolder = @"C:\Workbooks\Output";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get workbook files (any supported format) from the input folder
            string[] workbookFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
            int filesToProcess = Math.Min(10, workbookFiles.Length);

            // Freeze configuration: freeze the first row, no column freeze
            int freezeRow = 1;    // rows above this index are frozen
            int freezeColumn = 0; // columns left of this index are frozen

            for (int i = 0; i < filesToProcess; i++)
            {
                string inputPath = workbookFiles[i];

                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"File not found: {inputPath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(inputPath);

                    // Apply freeze panes to each worksheet
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Freeze first row (freezeRow rows, freezeColumn columns)
                        sheet.FreezePanes(freezeRow, freezeColumn, freezeRow, freezeColumn);
                    }

                    // Build the output file path (preserve original file name)
                    string fileName = Path.GetFileName(inputPath);
                    string outputPath = Path.Combine(outputFolder, fileName);

                    // Save the modified workbook
                    workbook.Save(outputPath);
                    Console.WriteLine($"Processed: {fileName}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{inputPath}': {ex.Message}");
                }
            }

            Console.WriteLine($"Processed {filesToProcess} workbook(s) and applied row freeze configuration.");
        }
    }
}
