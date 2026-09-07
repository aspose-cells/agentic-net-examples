// Title: Convert an Excel .xlsx file to UTF-8 encoded HTML using Aspose.Cells for .NET
// AI Prompts: Write a C# program that loads a workbook from a given .xlsx path, sets HtmlSaveOptions.Encoding to UTF-8, and saves it as an HTML file. | Show how to verify the existence of an input Excel file before exporting it to UTF-8 HTML with Aspose.Cells. | Demonstrate using Aspose.Cells HtmlSaveOptions to export only the active worksheet to a UTF-8 encoded HTML document.
// Common Searches: asp.net how to export Excel to HTML with UTF-8 encoding using Aspose.Cells | c# set HtmlSaveOptions.Encoding to UTF8 when saving workbook as HTML | check file exists before Aspose.Cells workbook.Save to HTML | convert .xlsx to UTF-8 HTML with Aspose.Cells .NET example
// Tags: Aspose.Cells HtmlSaveOptions UTF-8 encoding | C# export workbook to HTML Aspose.Cells | verify Excel file existence before Aspose.Cells conversion | save active worksheet as HTML Aspose.Cells | UTF-8 HTML output from .xlsx using Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Loads an .xlsx workbook, checks that the file exists, configures HtmlSaveOptions with Encoding = Encoding.UTF8, and saves the workbook as an HTML file using Aspose.Cells, with basic error handling.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.html";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    Encoding = Encoding.UTF8
                    // ExportActiveWorksheetOnly = true,
                    // ExportGridLines = true
                };

                // Save the workbook as an HTML file using the specified options
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully exported to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
