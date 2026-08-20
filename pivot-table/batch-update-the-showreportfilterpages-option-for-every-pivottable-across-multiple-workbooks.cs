// Title: Batch enable ShowReportFilterPage for all PivotTables in multiple Excel files with Aspose.Cells (C#)
// Description: C# sample that loads a collection of Excel workbooks using Aspose.Cells, iterates every worksheet and PivotTable, calls ShowReportFilterPage for each page field, and saves the modified files to a target folder while handling missing files and runtime errors.
// Keywords: Aspose.Cells C# pivot table batch update | ShowReportFilterPage programmatically | process multiple Excel workbooks | enable report filter pages | pivot table page fields Aspose | bulk modify Excel pivot settings
// Common Searches: how to set ShowReportFilterPage for all pivot tables in C# | batch update pivot table report filter pages Aspose.Cells | iterate pivot tables across many workbooks .NET | save modified Excel files after changing pivot settings | Aspose.Cells example for ShowReportFilterPage
// Developer Intent: Automatically turn on the ShowReportFilterPage option for every page field of each PivotTable in a set of workbooks and write the updated files to a chosen directory.
// Use Cases: Standardize report‑filter worksheets across a corporate library before distribution. | Prepare workbooks for publishing by ensuring each pivot table creates separate filter pages. | Automate pivot‑table configuration during a data‑migration or ETL process.
// AI Prompts: Create C# code with Aspose.Cells that enables ShowReportFilterPage for all pivot tables in a folder of Excel files and logs any errors. | Refactor the batch method to run asynchronously and report progress for large numbers of workbooks. | Explain the impact of ShowReportFilterPage on generated worksheets and show how to verify the changes via code.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsBatchUpdate
{
    // C# sample that loads a collection of Excel workbooks using Aspose.Cells, iterates every worksheet and PivotTable, calls ShowReportFilterPage for each page field, and saves the modified files to a target folder while handling missing files and runtime errors.
    public class PivotReportFilterUpdater
    {
        /// <param name="inputFiles">Full paths of the source workbooks.</param>
        /// <param name="outputDirectory">Directory where the modified workbooks will be saved.</param>
        public static void UpdateShowReportFilterPages(string[] inputFiles, string outputDirectory)
        {
            // Ensure the output directory exists
            if (!Directory.Exists(outputDirectory))
                Directory.CreateDirectory(outputDirectory);

            foreach (string inputFile in inputFiles)
            {
                // Verify the source file exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Source file not found: {inputFile}");
                    continue;
                }

                try
                {
                    // Load the workbook from the source file
                    Workbook workbook = new Workbook(inputFile);

                    // Iterate through all worksheets in the workbook
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Access the collection of pivot tables on the current worksheet
                        PivotTableCollection pivotTables = sheet.PivotTables;

                        // Process each pivot table
                        foreach (PivotTable pivotTable in pivotTables)
                        {
                            // For each page field in the pivot table, generate a separate report filter page
                            foreach (PivotField pageField in pivotTable.PageFields)
                            {
                                pivotTable.ShowReportFilterPage(pageField);
                            }
                        }
                    }

                    // Construct the output file path (preserve original file name)
                    string outputPath = Path.Combine(outputDirectory, Path.GetFileName(inputFile));

                    // Save the modified workbook
                    workbook.Save(outputPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{inputFile}': {ex.Message}");
                }
            }
        }

        // Example usage
        public static void Main()
        {
            // Define source workbook files (could be populated dynamically)
            string[] sourceFiles = new string[]
            {
                @"C:\Data\Workbook1.xlsx",
                @"C:\Data\Workbook2.xlsx",
                @"C:\Data\Workbook3.xlsx"
            };

            // Define where the updated workbooks should be saved
            string destinationFolder = @"C:\Data\UpdatedWorkbooks";

            // Perform the batch update
            UpdateShowReportFilterPages(sourceFiles, destinationFolder);

            Console.WriteLine("Batch update completed.");
        }
    }
}
