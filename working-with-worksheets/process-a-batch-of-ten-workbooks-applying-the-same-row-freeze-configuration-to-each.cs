// Title: Batch freeze the first row in multiple Excel workbooks using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that accepts an array of .xlsx file paths, opens each workbook with Aspose.Cells, freezes the top row on every worksheet, and saves the modified file with a "_frozen" suffix. | Create a C# console application that loops through ten Excel files, applies FreezePanes(0,0,1,0) to each sheet via Aspose.Cells, and logs missing files without interrupting the batch.
// Common Searches: c# aspose.cells batch freeze first row in all worksheets of multiple Excel files | how to apply FreezePanes to several workbooks programmatically with Aspose.Cells | process a list of .xlsx files and add row freeze using Aspose.Cells in .NET | save modified Excel workbooks with custom suffix after applying freeze panes in C#
// Tags: freeze panes first row aspose.cells c# | batch process multiple excel workbooks aspose.cells | save workbook with custom suffix aspose.cells | iterate worksheets apply freeze panes c#

using System;
using System.IO;
using Aspose.Cells;

// The program iterates over ten specified .xlsx files, loads each workbook with Aspose.Cells, freezes the first row (no columns) on every worksheet using FreezePanes, saves the result with a "_frozen" suffix, and logs progress while handling missing files.
class BatchFreezeRows
{
    static void Main()
    {
        // Define the row freeze configuration: freeze the first row, no columns frozen.
        int rowsToFreeze = 1;
        int columnsToFreeze = 0;

        // List of input workbook file paths (adjust paths as needed).
        string[] inputFiles = new string[10]
        {
            @"C:\Workbooks\Book1.xlsx",
            @"C:\Workbooks\Book2.xlsx",
            @"C:\Workbooks\Book3.xlsx",
            @"C:\Workbooks\Book4.xlsx",
            @"C:\Workbooks\Book5.xlsx",
            @"C:\Workbooks\Book6.xlsx",
            @"C:\Workbooks\Book7.xlsx",
            @"C:\Workbooks\Book8.xlsx",
            @"C:\Workbooks\Book9.xlsx",
            @"C:\Workbooks\Book10.xlsx"
        };

        try
        {
            // Process each workbook.
            foreach (string inputPath in inputFiles)
            {
                // Verify the input file exists.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"File not found: {inputPath}");
                    continue;
                }

                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);

                // Apply the freeze configuration to each worksheet.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // FreezePanes(row, column, totalRows, totalColumns)
                    // Row and column specify the split position (0‑based). To freeze the first row,
                    // split occurs before row 1 (row index 0). TotalRows indicates how many rows to freeze.
                    sheet.FreezePanes(0, 0, rowsToFreeze, columnsToFreeze);
                }

                // Build the output file path (adds "_frozen" before the extension).
                string directory = Path.GetDirectoryName(inputPath) ?? string.Empty;
                string filenameWithoutExt = Path.GetFileNameWithoutExtension(inputPath);
                string extension = Path.GetExtension(inputPath);
                string outputPath = Path.Combine(directory, $"{filenameWithoutExt}_frozen{extension}");

                // Ensure the output directory exists.
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Processed and saved: {outputPath}");
            }

            Console.WriteLine("Batch processing completed. Row freeze applied to all workbooks.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
