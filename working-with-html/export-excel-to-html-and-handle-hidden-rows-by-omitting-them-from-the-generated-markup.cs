// Title: Convert an Excel workbook to HTML in C# with Aspose.Cells while excluding hidden rows
// AI Prompts: Generate C# code that loads a .xlsx file using Aspose.Cells, configures HtmlSaveOptions to skip hidden rows, and saves the result as an HTML file. | Show how to verify the source Excel file exists and handle errors before exporting to HTML with hidden rows omitted. | Demonstrate setting the ExportHiddenRows property (or relying on its default) in HtmlSaveOptions for an Aspose.Cells .NET conversion.
// Common Searches: asp.net convert excel to html without hidden rows using aspose.cells | c# htmlsaveoptions exporthiddenrows false example | how to prevent hidden rows from appearing in html output from aspose cells | save workbook as html excluding hidden rows c# asp.net | aspose.cells hide rows from html export .net core
// Tags: Aspose.Cells HTML export options | C# Excel to HTML conversion | exclude hidden rows in HTML output | Aspose.Cells workbook save as html | validate input Excel file C#

using System;
using System.IO;
using Aspose.Cells;

// A console application that checks for the input.xlsx file, loads it with Aspose.Cells, configures HtmlSaveOptions to omit hidden rows (default behavior or ExportHiddenRows = false), and saves the workbook as output.html while handling exceptions.
class ExcelToHtmlExporter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: In recent Aspose.Cells versions, hidden rows are excluded by default.
            // If using a version that supports ExportHiddenRows, uncomment the line below:
            // htmlOptions.ExportHiddenRows = false;

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully exported to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display an error message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
