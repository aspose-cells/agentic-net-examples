// Title: Save an Excel workbook as HTML using Aspose.Cells default HtmlSaveOptions in C#
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells and calls Workbook.Save with SaveFormat.Html to produce an HTML file using the library's default options. | Show a minimal example of exporting a Workbook to HTML in C# without configuring any HtmlSaveOptions.
// Common Searches: aspnet convert xlsx to html with Aspose.Cells default settings | c# Aspose.Cells export workbook to html without custom options | how to use Workbook.Save to generate html from excel in C#
// Tags: Aspose.Cells save workbook as HTML C# | default HtmlSaveOptions export Excel to HTML | Workbook.Save with SaveFormat.Html example | convert .xlsx to .html using Aspose.Cells

using System;
using Aspose.Cells;

// The code loads 'input.xlsx' into an Aspose.Cells Workbook and saves it as 'output.html' by invoking Workbook.Save with SaveFormat.Html, which applies the default HtmlSaveOptions automatically.
class Program
{
    static void Main()
    {
        // Load the Excel workbook from a file (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Save the workbook as HTML using the default HtmlSaveOptions
        // The Save method automatically applies default options when only format is specified
        workbook.Save("output.html", SaveFormat.Html);
    }
}
