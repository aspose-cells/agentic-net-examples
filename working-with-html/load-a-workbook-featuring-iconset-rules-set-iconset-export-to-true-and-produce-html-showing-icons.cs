// Title: Export an Excel workbook with IconSet conditional formatting to HTML while preserving icons using Aspose.Cells for .NET
// AI Prompts: Write a C# program that opens an .xlsx workbook containing IconSet conditional formatting and saves it as an HTML file where the icons are rendered, using Aspose.Cells. | Demonstrate how to configure Aspose.Cells HtmlSaveOptions to include IconSet icons during HTML conversion. | Create error‑handling code that verifies the source Excel file exists before exporting it to HTML with IconSet icons using Aspose.Cells.
// Common Searches: asp.net export excel file containing icons set conditional formatting to html with Aspose.Cells | preserve conditional formatting icons when converting workbook to html in c# | how to use HtmlSaveOptions to keep icon set rules in html output | c# example for exporting excel with icon sets to html using aspose.cells
// Tags: Aspose.Cells HtmlSaveOptions IconSet export | C# Excel to HTML conversion preserving conditional formatting icons | IconSet conditional formatting HTML output .NET | handle missing workbook file Aspose.Cells | export workbook with icons to html using Aspose

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook that contains IconSet conditional formatting, applies HtmlSaveOptions, and saves it as an HTML file where the icons are rendered, with basic file‑existence error handling.
class IconSetHtmlExport
{
    static void Main()
    {
        try
        {
            const string inputFile = "InputWithIconSet.xlsx";
            const string outputFile = "IconSetExported.html";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" was not found.");
                return;
            }

            // Load the workbook that contains IconSet conditional formatting rules
            Workbook workbook = new Workbook(inputFile);

            // Configure HTML save options (conditional formatting is exported by default)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as HTML; the resulting file will display the icons
            workbook.Save(outputFile, htmlOptions);

            Console.WriteLine($"Workbook successfully exported to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
