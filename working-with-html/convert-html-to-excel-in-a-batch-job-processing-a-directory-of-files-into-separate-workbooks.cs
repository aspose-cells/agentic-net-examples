// Title: C# Batch Convert HTML Files to Individual Excel Workbooks with Aspose.Cells
// Description: A console utility that validates input and output folders, scans a directory for *.html files, loads each file as an Aspose.Cells Workbook using LoadFormat.Html, and saves it as a separate .xlsx workbook with the same base name while logging successes and errors.
// Keywords: Aspose.Cells HTML to Excel batch | C# convert multiple HTML files | directory processing Aspose.Cells | load HTML workbook C# | save workbook as Xlsx programmatically
// Common Searches: batch convert html to xlsx c# | process folder of html files Aspose.Cells | convert all html tables to separate excel files | automate html to excel conversion .NET
// Developer Intent: Transform every HTML file in a specified folder into its own Excel workbook using Aspose.Cells for .NET.
// Use Cases: Nightly automation that turns exported HTML reports into Excel sheets for analytics pipelines. | Migration of a legacy HTML data archive into individual .xlsx files for business users. | Command‑line tool that processes bulk HTML invoices and generates separate Excel files for accounting.
// AI Prompts: Add detailed file‑level logging to the batch conversion and write errors to a log file. | Modify the program to accept input and output directories as command‑line arguments instead of hard‑coded paths. | Show how to implement a console progress bar while converting a large set of HTML files with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace HtmlToExcelBatch
{
    // A console utility that validates input and output folders, scans a directory for *.html files, loads each file as an Aspose.Cells Workbook using LoadFormat.Html, and saves it as a separate .xlsx workbook with the same base name while logging successes and errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Directory containing HTML files
            string inputDirectory = @"C:\InputHtmlFiles";

            // Directory where converted Excel files will be saved
            string outputDirectory = @"C:\OutputExcelFiles";

            try
            {
                // Verify input directory exists
                if (!Directory.Exists(inputDirectory))
                {
                    Console.WriteLine($"Input directory does not exist: {inputDirectory}");
                    return;
                }

                // Ensure the output directory exists
                if (!Directory.Exists(outputDirectory))
                    Directory.CreateDirectory(outputDirectory);

                // Load options to treat source files as HTML
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);

                // Process each .html file in the input directory
                foreach (string htmlFilePath in Directory.GetFiles(inputDirectory, "*.html"))
                {
                    try
                    {
                        // Load the HTML file into a workbook
                        Workbook workbook = new Workbook(htmlFilePath, loadOptions);

                        // Determine output file name (same base name with .xlsx extension)
                        string fileNameWithoutExt = Path.GetFileNameWithoutExtension(htmlFilePath);
                        string excelFilePath = Path.Combine(outputDirectory, fileNameWithoutExt + ".xlsx");

                        // Save the workbook as an Excel file
                        workbook.Save(excelFilePath, SaveFormat.Xlsx);

                        Console.WriteLine($"Converted: {htmlFilePath} -> {excelFilePath}");
                    }
                    catch (Exception ex)
                    {
                        // Log any errors but continue processing other files
                        Console.WriteLine($"Error converting '{htmlFilePath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch conversion completed.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
