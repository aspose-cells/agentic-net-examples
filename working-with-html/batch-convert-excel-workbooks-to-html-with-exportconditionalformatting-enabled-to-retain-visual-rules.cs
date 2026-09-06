// Title: Batch convert multiple .xlsx workbooks to HTML with conditional formatting retained using Aspose.Cells for .NET
// AI Prompts: Generate a C# console program that scans a folder for all .xlsx files and saves each workbook as an .html file using Aspose.Cells, ensuring conditional formatting rules are exported. | Write C# code that creates an Aspose.Cells HtmlSaveOptions instance with ExportConditionalFormatting enabled and applies it while batch‑processing Excel files to HTML. | Provide a C# script that iterates through Excel files in a directory, loads each with Aspose.Cells, and outputs HTML files that preserve visual conditional formatting.
// Common Searches: aspocells batch convert xlsx to html preserving conditional formatting c# | how to export Excel conditional formatting to HTML using Aspose.Cells .NET | C# program to convert multiple Excel workbooks to HTML with visual rules retained | HtmlSaveOptions ExportConditionalFormatting example for bulk conversion | convert folder of .xlsx files to .html with Aspose.Cells and keep formatting
// Tags: batch xlsx to html conversion Aspose.Cells | export conditional formatting HtmlSaveOptions | Aspose.Cells HTML export preserving visual rules | C# process multiple workbooks to HTML | automated Excel to HTML conversion .NET

using System;
using System.IO;
using Aspose.Cells;

// The example defines input and output directories, enumerates all .xlsx files in the source folder, loads each workbook with Aspose.Cells, configures HtmlSaveOptions (which export conditional formatting by default), saves each workbook as an .html file in the target folder, logs progress, handles errors, and reports completion.
class Program
{
    static void Main(string[] args)
    {
        // Define the folder containing the Excel workbooks
        string inputFolder = @"C:\InputExcels";

        // Define the folder where the HTML files will be saved
        string outputFolder = @"C:\OutputHtml";

        // Ensure the input directory exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Retrieve all Excel files (adjust the pattern for .xls if needed)
        string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx");

        foreach (string excelPath in excelFiles)
        {
            // Verify the file exists before attempting to load
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"File not found: {excelPath}");
                continue;
            }

            try
            {
                // Load the Excel workbook
                Workbook workbook = new Workbook(excelPath);

                // Configure HTML save options (conditional formatting is exported by default)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

                // Build the output HTML file path
                string fileName = Path.GetFileNameWithoutExtension(excelPath);
                string htmlPath = Path.Combine(outputFolder, fileName + ".html");

                // Save the workbook as HTML using the configured options
                workbook.Save(htmlPath, htmlOptions);

                Console.WriteLine($"Converted: {excelPath} -> {htmlPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{excelPath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion to HTML completed.");
    }
}
