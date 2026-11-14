// Title: How to combine all worksheets of an Excel workbook into a single CSV file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export every worksheet to a single CSV file | combine multiple Excel sheets into one CSV using TxtSaveOptions | set ExportAllSheets = true when saving workbook as CSV in .NET | save workbook with all sheets concatenated to CSV with Aspose.Cells | C# code to export all sheets from .xlsx to combined .csv
// Tags: ExportAllSheets TxtSaveOptions CSV Aspose.Cells | save workbook as single CSV multiple worksheets | combine Excel sheets into one CSV .NET | Aspose.Cells CSV export all sheets | C# TxtSaveOptions SaveFormat.Csv configuration

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExportAllSheetsToCsv
{
    // The example loads an Excel workbook, configures TxtSaveOptions with SaveFormat.Csv, enables ExportAllSheets, sets a comma separator, and saves all worksheets concatenated into a single CSV file, handling missing files and runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook (can be .xlsx, .xls, etc.)
            string sourcePath = "input.xlsx";

            // Path for the combined CSV output
            string outputPath = "combined_output.csv";

            try
            {
                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Error: The source file \"{sourcePath}\" was not found.");
                    return;
                }

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(sourcePath);

                // Configure TXT save options to export all worksheets into a single CSV file
                TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    ExportAllSheets = true,
                    Separator = ','
                };

                // Save the workbook as CSV; all sheets will be concatenated into one file
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine("All worksheets have been exported to a single CSV file successfully.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
