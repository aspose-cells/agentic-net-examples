// Title: Batch process 100 Excel workbooks in C# with Aspose.Cells: add author and custom properties, remove unused styles
// AI Prompts: Create a C# console application that loops from Workbook1.xlsx to Workbook100.xlsx, loads each file with Aspose.Cells, sets the built‑in Author property, adds a custom ProcessedDate property, invokes RemoveUnusedStyles, and saves the modified workbook to a designated output folder. | Enhance the batch processor to accept command‑line parameters for input folder, output folder, and the start/end workbook numbers, and write a log file that records missing source files and any processing exceptions.
// Common Searches: C# Aspose.Cells loop through multiple XLSX files and set document properties | remove unused styles from a batch of Excel workbooks using Aspose.Cells | add custom ProcessedDate property to many Excel files programmatically .NET | how to handle missing Excel files in a bulk processing script with Aspose.Cells | save modified workbooks to a different directory using Aspose.Cells C#
// Tags: batch update built‑in author property Aspose.Cells | add custom ProcessedDate property to XLSX files | clean unused cell styles Aspose.Cells | bulk workbook processing C# console | output directory saving Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program iterates over 100 Excel files, loads each with Aspose.Cells, sets the Author built‑in property, adds a custom ProcessedDate property, removes any styles that are not referenced, and saves the updated workbook to an output folder while handling missing files and logging errors.
class BatchWorkbookProcessor
{
    static void Main()
    {
        // Define input and output directories
        string inputDir = @"C:\Workbooks\Input";
        string outputDir = @"C:\Workbooks\Output";

        // Ensure output directory exists
        Directory.CreateDirectory(outputDir);

        // Process workbooks named Workbook1.xlsx to Workbook100.xlsx
        for (int i = 1; i <= 100; i++)
        {
            string inputFile = Path.Combine(inputDir, $"Workbook{i}.xlsx");
            string outputFile = Path.Combine(outputDir, $"Workbook{i}.xlsx");

            // Skip if the input file does not exist
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                continue;
            }

            try
            {
                // Load the workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(inputFile);

                // ---- Optional Metadata Section ----
                // Set built‑in property (e.g., Author)
                workbook.BuiltInDocumentProperties["Author"].Value = "BatchProcessor";

                // Add custom property (e.g., ProcessedDate)
                workbook.CustomDocumentProperties.Add("ProcessedDate", DateTime.Now);

                // ---- Remove Unused Styles ----
                // Cleans up any style definitions that are not referenced by cells
                workbook.RemoveUnusedStyles();

                // Save the modified workbook (lifecycle rule: save)
                workbook.Save(outputFile, SaveFormat.Xlsx);
                Console.WriteLine($"Processed and saved: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file {inputFile}: {ex.Message}");
            }
        }
    }
}
