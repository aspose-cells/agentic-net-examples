// Title: Create a C# console app to convert an Excel file to HTML5 with fit‑to‑page width and Base64‑embedded images using Aspose.Cells
// AI Prompts: Generate a C# console program that accepts an Excel file path argument, sets each worksheet’s PageSetup.FitToPagesWide to 1, and saves the workbook as an HTML5 file with images embedded as Base64 using Aspose.Cells. | Demonstrate how to configure Aspose.Cells HtmlSaveOptions to export all worksheets, embed images as Base64, and produce HTML5 output in a .NET console application. | Add robust error handling for missing command‑line arguments, nonexistent files, and runtime exceptions during the Aspose.Cells HTML export.
// Common Searches: how to export excel to html5 with fit to page width using aspose.cells c# | c# console application convert workbook to html with base64 images asp.net | set FitToPagesWide for all worksheets aspose.cells example | htmlsaveoptions embed images as base64 aspose.cells c# | asp.net command line tool convert xlsx to html5 using aspose.cells
// Tags: aspocells htmlsaveoptions base64 images | excel to html5 conversion aspocells c# | fit-to-page width worksheet page setup aspocells | c# console utility excel html export | aspocells export all worksheets html

using System;
using System.IO;
using Aspose.Cells;

// A C# console utility that loads an Excel workbook from a command‑line path, applies a fit‑to‑page‑wide setting to every worksheet, configures HtmlSaveOptions to generate HTML5 with all sheets and Base64‑encoded images, and writes the result to a .html file while handling missing arguments and file errors.
class Program
{
    static void Main(string[] args)
    {
        // Verify that a file path was provided.
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: AsposeCellsHtmlExport <excel-file-path>");
            return;
        }

        string excelPath = args[0];

        // Ensure the input file exists to avoid FileNotFoundException.
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Error: The file \"{excelPath}\" does not exist.");
            return;
        }

        try
        {
            // Load the Excel workbook.
            Workbook workbook = new Workbook(excelPath);

            // Apply fit‑to‑page settings for each worksheet.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Fit the worksheet width to one page; height is unlimited.
                sheet.PageSetup.FitToPagesWide = 1;
                sheet.PageSetup.FitToPagesTall = 0;
            }

            // Configure HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export all worksheets (set to true for a single sheet if needed).
                ExportActiveWorksheetOnly = false,
                // Embed images directly into the HTML as Base64 strings.
                ExportImagesAsBase64 = true,
                // Use HTML5 for better compatibility.
                HtmlVersion = HtmlVersion.Html5
            };

            // Determine the output HTML file path.
            string htmlPath = Path.ChangeExtension(excelPath, ".html");

            // Save the workbook as an HTML file using the configured options.
            workbook.Save(htmlPath, htmlOptions);

            Console.WriteLine($"HTML file successfully created at: {htmlPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
