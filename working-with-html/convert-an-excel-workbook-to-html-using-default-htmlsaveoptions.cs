// Title: How to save an Excel workbook as HTML using Aspose.Cells default HtmlSaveOptions in C#
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells and saves it as an HTML file using the default HtmlSaveOptions. | Provide a minimal example showing how to export a Workbook to HTML without configuring any HtmlSaveOptions in Aspose.Cells for .NET. | Write a C# snippet that demonstrates converting input.xlsx to output.html using Aspose.Cells default HTML save settings.
// Common Searches: Aspose.Cells C# export workbook to HTML using default settings | save Excel workbook as HTML file with Aspose.Cells without custom options | C# example converting .xlsx to .html using Aspose.Cells default configuration | how to use Aspose.Cells HtmlSaveOptions to generate HTML from an Excel file
// Tags: Aspose.Cells default HTML export | C# convert Excel to HTML Aspose.Cells | Workbook.Save HTML Aspose.Cells | Excel to HTML conversion .NET Aspose | Aspose.Cells HTML output without customization

using System;
using Aspose.Cells;

// Loads 'input.xlsx' into an Aspose.Cells Workbook, creates a default HtmlSaveOptions instance, and saves the workbook as 'output.html' using the built‑in HTML export settings.
class Program
{
    static void Main()
    {
        // Load the Excel workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Create default HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Save the workbook as an HTML file using the default options
        workbook.Save("output.html", htmlOptions);
    }
}
