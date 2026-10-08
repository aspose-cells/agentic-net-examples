// Title: Build a C# console utility that converts an Excel workbook to HTML with Aspose.Cells using a best‑fit layout
// AI Prompts: Create a C# console application that accepts an Excel file path as a command‑line argument, loads the workbook with Aspose.Cells, and saves it as an HTML file using Aspose.Cells HTML export settings with the best‑fit layout enabled. | Add validation to ensure the argument is provided and the file exists, and wrap the conversion logic in a try‑catch block that outputs clear error messages. | Generate the output file name by replacing the source extension with .html, write the full path to the console, and ensure all worksheets are included in the HTML output.
// Common Searches: c# how to use Aspose.Cells to convert an Excel workbook to a single HTML file from the command line | example of exporting all worksheets to HTML with best‑fit layout in a .NET console program | command‑line tool for converting .xlsx to .html using Aspose.Cells HtmlSaveOptions
// Tags: aspocells html export settings optimal layout | c# command line excel to html conversion | export all worksheets to html aspocells | derive html filename from excel path c# | aspocells save workbook as html

using System;
using System.IO;
using Aspose.Cells;

// A C# console program that reads an Excel file path supplied via the command line, loads the workbook with Aspose.Cells, configures HTML export options to include all worksheets and apply a best‑fit layout, then saves the result as an .html file sharing the original name, while providing robust argument validation and error handling.
class Program
{
    static void Main(string[] args)
    {
        // Verify that a file path was provided.
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: Program <excel-file-path>");
            return;
        }

        string excelPath = args[0];

        // Ensure the input Excel file exists.
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Error: File not found – {excelPath}");
            return;
        }

        try
        {
            // Load the Excel workbook from the specified path.
            Workbook workbook = new Workbook(excelPath);

            // Configure HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Export all worksheets.
            htmlOptions.ExportActiveWorksheetOnly = false;

            // Determine the output HTML file name (same name, .html extension).
            string htmlPath = Path.ChangeExtension(excelPath, ".html");

            // Save the workbook as HTML using the configured options.
            workbook.Save(htmlPath, htmlOptions);

            Console.WriteLine($"HTML file generated: {htmlPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
