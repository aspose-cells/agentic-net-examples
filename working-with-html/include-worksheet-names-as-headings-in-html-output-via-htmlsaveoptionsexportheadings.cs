// Title: Export worksheet names as headings when saving an Excel workbook to HTML with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# export worksheet name as HTML heading | How to save Excel as HTML with sheet titles using HtmlSaveOptions | C# example for HtmlSaveOptions ExportRowColumnHeadings property | Convert .xlsx to .html and include worksheet headings Aspose.Cells | Include sheet names as headings in HTML output from Excel C#
// Tags: Aspose.Cells HtmlSaveOptions ExportRowColumnHeadings | C# Excel to HTML conversion with sheet headings | save workbook as HTML with worksheet titles | HTML export of Excel including sheet names

using System;
using System.IO;
using Aspose.Cells;

// The code verifies the input XLSX file, loads it into an Aspose.Cells Workbook, sets HtmlSaveOptions.ExportRowColumnHeadings to true so each worksheet name appears as a heading, and saves the workbook as an HTML file.
class Program
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

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options to export worksheet names as headings
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Use the updated property name (ExportHeadings is obsolete)
                ExportRowColumnHeadings = true
            };

            // Save the workbook as an HTML file using the options overload
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
