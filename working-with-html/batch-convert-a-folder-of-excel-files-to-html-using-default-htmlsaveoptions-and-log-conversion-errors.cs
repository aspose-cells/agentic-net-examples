// Title: Batch convert Excel files to HTML with Aspose.Cells (C#) and log errors
// Description: A C# console utility that scans a given directory for Excel workbooks (.xls, .xlsx, .xlsm, .xlsb, .csv, .ods), converts each file to HTML using Aspose.Cells ConversionUtility with default HtmlSaveOptions, and records any conversion exceptions without stopping the batch run.
// Keywords: Aspose.Cells batch conversion | Excel to HTML C# | ConversionUtility default options | process multiple Excel formats | log conversion errors | folder based Excel conversion | C# .NET Aspose.Cells example
// Common Searches: convert all Excel files in a folder to HTML Aspose.Cells | C# batch Excel to HTML code sample | Aspose.Cells ConversionUtility error handling | default HtmlSaveOptions for Excel conversion | how to log failures during batch Excel conversion
// Developer Intent: Automatically transform every supported Excel workbook in a specified folder into an HTML file using Aspose.Cells while capturing and reporting any conversion failures.
// Use Cases: Generate web‑ready versions of a large spreadsheet library for intranet publishing. | Create a scheduled command‑line tool that produces HTML reports from nightly‑updated Excel data. | Integrate into CI/CD pipelines to verify that all Excel assets render correctly as HTML without breaking the build.
// AI Prompts: Write C# code that iterates over a directory, converts each .xls/.xlsx/.xlsm/.xlsb/.csv/.ods file to HTML with Aspose.Cells default HtmlSaveOptions, and logs exceptions to the console. | Show an error‑handling pattern for a batch Excel‑to‑HTML conversion loop that continues processing after a failure and writes detailed messages to a log file. | Explain how to modify the sample to output HTML files to a separate folder while preserving original filenames and maintaining the same error‑logging behavior.

using System;
using System.IO;
using System.Linq;
using Aspose.Cells.Utility;

namespace BatchExcelToHtml
{
    // A C# console utility that scans a given directory for Excel workbooks (.xls, .xlsx, .xlsm, .xlsb, .csv, .ods), converts each file to HTML using Aspose.Cells ConversionUtility with default HtmlSaveOptions, and records any conversion exceptions without stopping the batch run.
    class Program
    {
        static void Main(string[] args)
        {
            // Specify the folder containing Excel files.
            // You can pass the folder path as a command‑line argument or set it directly here.
            string sourceFolder = args.Length > 0 ? args[0] : @"C:\ExcelFiles";

            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Folder does not exist: {sourceFolder}");
                return;
            }

            // Define the Excel file extensions to process.
            string[] excelExtensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".csv", ".ods" };

            // Get all files with the supported extensions.
            var excelFiles = Directory.GetFiles(sourceFolder)
                                      .Where(f => excelExtensions.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase));

            foreach (var sourcePath in excelFiles)
            {
                try
                {
                    // Destination HTML file path – same name with .html extension.
                    string destPath = Path.ChangeExtension(sourcePath, ".html");

                    // Convert using Aspose.Cells ConversionUtility with default options.
                    ConversionUtility.Convert(sourcePath, destPath);

                    Console.WriteLine($"Converted: {Path.GetFileName(sourcePath)} -> {Path.GetFileName(destPath)}");
                }
                catch (Exception ex)
                {
                    // Log conversion errors without stopping the batch process.
                    Console.WriteLine($"Error converting '{Path.GetFileName(sourcePath)}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
