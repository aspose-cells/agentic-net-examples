// Title: Batch enable ShowReportFilterPages for every PivotTable in multiple Excel workbooks using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a list of Excel file paths, iterates through each worksheet and its PivotTable collection, calls ShowReportFilterPage for every page field, and saves the workbook to a specified output directory. | Add logging to the sample so that workbook names without any PivotTables are recorded while performing the batch ShowReportFilterPages update. | Refactor the program to accept the input file list and the output folder as command‑line arguments instead of hard‑coded values.
// Common Searches: Aspose.Cells how to turn on ShowReportFilterPages for all pivot tables in a batch of workbooks | C# batch processing of Excel files to update pivot table report filter pages | iterate through worksheets and pivot tables with Aspose.Cells to call ShowReportFilterPage | automate pivot table page field settings across multiple .xlsx files using .NET | example code for updating ShowReportFilterPage option in several workbooks with Aspose.Cells
// Tags: batch update pivot table ShowReportFilterPages Aspose.Cells | iterate worksheets pivot tables C# Aspose | process multiple Excel workbooks Aspose.Cells | automate pivot table page field settings .NET | save updated workbooks with suffix Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotReportFilterBatchUpdate
{
    // The program loads each workbook from a supplied list, walks through every worksheet and its PivotTable collection, invokes ShowReportFilterPage for each page field to display report filter pages, and saves the modified workbook with an "_Updated" suffix to a target folder.
    public class Program
    {
        // Entry point
        public static void Main(string[] args)
        {
            // Example input: array of workbook file paths to process
            string[] inputFiles = new string[]
            {
                @"C:\Data\Workbook1.xlsx",
                @"C:\Data\Workbook2.xlsx",
                @"C:\Data\Workbook3.xlsx"
            };

            // Folder where updated workbooks will be saved
            string outputFolder = @"C:\Data\Updated";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Process each workbook
            foreach (string inputPath in inputFiles)
            {
                // Verify the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"File not found: {inputPath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(inputPath);

                    // Iterate through all worksheets
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Access the collection of pivot tables on the worksheet
                        PivotTableCollection pivots = sheet.PivotTables;

                        // Process each pivot table
                        foreach (PivotTable pivot in pivots)
                        {
                            // Show report filter pages for every page field in the pivot table
                            foreach (PivotField pageField in pivot.PageFields)
                            {
                                // ShowReportFilterPage method
                                pivot.ShowReportFilterPage(pageField);
                            }
                        }
                    }

                    // Build output file path
                    string fileName = Path.GetFileNameWithoutExtension(inputPath);
                    string outputPath = Path.Combine(outputFolder, $"{fileName}_Updated.xlsx");

                    // Save the modified workbook
                    workbook.Save(outputPath);
                }
                catch (Exception ex)
                {
                    // Log any errors and continue with the next file
                    Console.WriteLine($"Error processing '{inputPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch update completed.");
        }
    }
}
