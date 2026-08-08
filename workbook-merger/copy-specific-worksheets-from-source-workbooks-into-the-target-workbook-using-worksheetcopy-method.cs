// Title: Copy specific worksheets from multiple Excel files into one workbook with Aspose.Cells for .NET (C#)
// Description: The sample defines an array of source workbook paths and a parallel array of worksheet names to copy, creates an empty target workbook, iterates through each source file, loads it, validates the presence of each requested sheet, adds a new sheet with the same name to the target, transfers data and formatting via Worksheet.Copy, and finally saves the merged file as CombinedWorkbook.xlsx while handling missing files or sheets gracefully.
// Keywords: Aspose.Cells | Worksheet.Copy | C# Excel merge | copy selected sheets | combine workbooks | merge specific worksheets | Excel file consolidation .NET | Aspose.Cells example | programmatic worksheet copy | Excel sheet merging
// Common Searches: Aspose.Cells copy specific worksheets from multiple workbooks | C# example for merging selected Excel sheets | How to use Worksheet.Copy to combine Excel files | Copy sheets with formatting using Aspose.Cells .NET | Merge Excel workbooks by selected sheets C#
// Developer Intent: Merge chosen worksheets from several source workbooks into a single target workbook using Aspose.Cells.
// Use Cases: Consolidate monthly department reports (Data, Summary) into a master workbook for executive review. | Extract "Report" sheets from project files and assemble them into a unified analysis workbook. | Create a master data set by pulling specific sheets from multiple regional Excel files. | Automate the generation of a combined financial statement by copying designated worksheets from quarterly files.
// AI Prompts: Generate a C# snippet that copies a list of worksheet names from multiple Excel files into one workbook with Aspose.Cells, including checks for missing files and sheets. | Explain how Worksheet.Copy retains cell formulas, styles, and page settings when merging selected sheets from different workbooks. | Suggest performance optimizations and more robust error handling for the provided worksheet‑copying code.

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetCopyExample
{
    // The sample defines an array of source workbook paths and a parallel array of worksheet names to copy, creates an empty target workbook, iterates through each source file, loads it, validates the presence of each requested sheet, adds a new sheet with the same name to the target, transfers data and formatting via Worksheet.Copy, and finally saves the merged file as CombinedWorkbook.xlsx while handling missing files or sheets gracefully.
    class Program
    {
        static void Main()
        {
            try
            {
                // Paths to source workbooks
                string[] sourceFiles = { "Source1.xlsx", "Source2.xlsx" };

                // For each source workbook define the worksheet names to copy
                string[][] sheetsToCopy = {
                    new string[] { "Data", "Summary" },   // from Source1.xlsx
                    new string[] { "Report" }            // from Source2.xlsx
                };

                // Create an empty target workbook and remove the default sheet
                Workbook targetWorkbook = new Workbook();
                targetWorkbook.Worksheets.Clear();

                // Iterate over each source workbook
                for (int i = 0; i < sourceFiles.Length; i++)
                {
                    string sourcePath = sourceFiles[i];

                    // Verify source file exists
                    if (!File.Exists(sourcePath))
                    {
                        Console.WriteLine($"Source file not found: {sourcePath}");
                        continue; // Skip to next source workbook
                    }

                    try
                    {
                        // Load the source workbook
                        Workbook sourceWorkbook = new Workbook(sourcePath);

                        // Copy each specified worksheet into the target workbook
                        foreach (string sheetName in sheetsToCopy[i])
                        {
                            // Ensure the worksheet exists in the source workbook
                            Worksheet sourceSheet = sourceWorkbook.Worksheets[sheetName];
                            if (sourceSheet == null)
                            {
                                Console.WriteLine($"Worksheet '{sheetName}' not found in '{sourcePath}'.");
                                continue;
                            }

                            // Add a new worksheet to the target workbook with the same name
                            Worksheet targetSheet = targetWorkbook.Worksheets.Add(sheetName);

                            // Copy the contents and formatting from the source worksheet
                            targetSheet.Copy(sourceSheet);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing '{sourcePath}': {ex.Message}");
                    }
                }

                // Save the combined workbook
                string outputPath = "CombinedWorkbook.xlsx";
                targetWorkbook.Save(outputPath);
                Console.WriteLine($"Worksheets copied successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
