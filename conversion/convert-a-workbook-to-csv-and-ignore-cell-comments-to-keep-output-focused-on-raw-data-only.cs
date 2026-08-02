// Title: Convert Excel Workbook to CSV without Comments using Aspose.Cells for .NET (C#)
// Description: Loads an .xlsx file, clears all worksheet comments, applies TxtSaveOptions (CSV format with trimmed leading blank rows and columns), and saves a clean CSV. Includes file‑existence verification and exception handling.
// Keywords: Aspose.Cells CSV export C# | remove Excel comments Aspose | TxtSaveOptions CSV | trim leading blank rows Aspose.Cells | C# workbook to CSV | clear worksheet comments | Aspose.Cells SaveFormat.Csv | Excel to CSV without comments | Aspose.Cells .NET conversion | CSV export ignoring comments
// Common Searches: Aspose.Cells export Excel to CSV without comments | C# remove cell comments before CSV conversion Aspose | How to clear comments in workbook using Aspose.Cells | Save workbook as CSV trimming empty rows Aspose.Cells | Convert .xlsx to .csv in .NET ignoring comments
// Developer Intent: Generate a CSV file from an Excel workbook while stripping all cell comments to keep only raw data.
// Use Cases: Produce clean CSV reports from user‑uploaded spreadsheets for analytics pipelines. | Prepare data extracts for systems that cannot interpret Excel comments, reducing payload size. | Automate nightly batch jobs that convert multiple workbooks to CSV after removing comments. | Create CSV files for import into databases where comments would cause parsing errors.
// AI Prompts: Write C# code using Aspose.Cells to convert an .xlsx to .csv, ensuring all comments are removed and leading blank rows/columns are trimmed. | Explain the effect of TxtSaveOptions.TrimLeadingBlankRowAndColumn on CSV output and when to use it. | Suggest performance‑optimized ways to clear comments in large workbooks before exporting to CSV with Aspose.Cells. | Provide troubleshooting steps if the CSV file still contains comment text after using ClearComments.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvExport
{
    // Loads an .xlsx file, clears all worksheet comments, applies TxtSaveOptions (CSV format with trimmed leading blank rows and columns), and saves a clean CSV. Includes file‑existence verification and exception handling.
    public class ExportWorkbookToCsv
    {
        public static void Run()
        {
            string sourcePath = "input.xlsx";
            string outputPath = "output.csv";

            try
            {
                // Verify source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the workbook from the file
                Workbook workbook = new Workbook(sourcePath);

                // Remove all comments from each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.ClearComments();
                }

                // Configure CSV save options
                TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    TrimLeadingBlankRowAndColumn = true
                };

                // Save the workbook as CSV
                workbook.Save(outputPath, csvOptions);
                Console.WriteLine($"Workbook has been exported to CSV without comments: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during export: {ex.Message}");
            }
        }

        // Entry point required for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
