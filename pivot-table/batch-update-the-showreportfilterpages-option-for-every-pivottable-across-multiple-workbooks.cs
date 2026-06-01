using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotReportFilterBatchUpdate
{
    class Program
    {
        static void Main(string[] args)
        {
            // List of workbook file paths to process
            string[] inputFiles = new string[]
            {
                @"C:\Data\Workbook1.xlsx",
                @"C:\Data\Workbook2.xlsx",
                // Add more file paths as needed
            };

            // Folder where the updated workbooks will be saved
            string outputFolder = @"C:\Data\Updated";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Process each workbook
            foreach (string inputPath in inputFiles)
            {
                try
                {
                    // Verify the input file exists before loading
                    if (!File.Exists(inputPath))
                    {
                        Console.WriteLine($"File not found: {inputPath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(inputPath);

                    // Iterate through all worksheets
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Iterate through all pivot tables in the worksheet
                        foreach (PivotTable pivotTable in sheet.PivotTables)
                        {
                            // For each page field, generate a separate report filter page
                            foreach (PivotField pageField in pivotTable.PageFields)
                            {
                                pivotTable.ShowReportFilterPage(pageField);
                            }
                        }
                    }

                    // Build the output file path (same file name, different folder)
                    string outputPath = Path.Combine(outputFolder, Path.GetFileName(inputPath));

                    // Save the modified workbook
                    workbook.Save(outputPath);

                    Console.WriteLine($"Processed and saved: {outputPath}");
                }
                catch (Exception ex)
                {
                    // Log any unexpected errors for this file and continue with the next one
                    Console.WriteLine($"Error processing '{inputPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch update completed.");
        }
    }
}