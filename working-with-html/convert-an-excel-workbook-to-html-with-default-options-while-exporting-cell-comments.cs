// Title: Convert an Excel workbook to HTML with default settings and include cell comments using Aspose.Cells for .NET (C#)
// Common Searches: c# convert xlsx to html with cell comments using aspose.cells | how to export excel cell comments to html with aspose.cells in .net | asp | asp | asp
// Tags: Aspose.Cells HTML conversion with cell comments | C# HtmlSaveOptions default export | Workbook.Save HTML Aspose.Cells | XLSX to HTML conversion default settings | export Excel comments to HTML .NET

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelToHtmlWithComments
{
    // The example loads an existing XLSX file into an Aspose.Cells Workbook, creates a default HtmlSaveOptions object (which automatically includes cell comments), and saves the workbook as an HTML file. It also checks for the source file's existence and handles any runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for source Excel file and generated HTML file
            string excelPath = "input.xlsx";
            string htmlPath = "output.html";

            try
            {
                // Verify that the input Excel file exists
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Error: Input file not found – {excelPath}");
                    return;
                }

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(excelPath);

                // Configure HTML save options (comments are exported by default)
                HtmlSaveOptions saveOptions = new HtmlSaveOptions();

                // Save the workbook as HTML using the configured options
                workbook.Save(htmlPath, saveOptions);

                Console.WriteLine($"Workbook has been converted to HTML with comments exported: {htmlPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
