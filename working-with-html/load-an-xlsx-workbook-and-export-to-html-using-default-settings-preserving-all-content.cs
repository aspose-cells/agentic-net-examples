// Title: Export an XLSX workbook to HTML with Aspose.Cells for .NET using default options
// Common Searches: how to export an Excel .xlsx file to HTML in C# using Aspose.Cells | Aspose.Cells C# save workbook as HTML with default HtmlSaveOptions | convert XLSX to HTML preserving formulas and formatting with Aspose.Cells .NET | sample code for saving Excel workbook to HTML using Aspose.Cells default settings
// Tags: Aspose.Cells export XLSX to HTML | HtmlSaveOptions default configuration | C# save workbook as HTML | preserve Excel formatting in HTML conversion | Aspose.Cells .NET HTML export

using System;
using Aspose.Cells;

// Loads an XLSX workbook from a file and saves it as an HTML document using Aspose.Cells for .NET with the default HtmlSaveOptions, ensuring all worksheet content and formatting are retained.
class Program
{
    static void Main()
    {
        // Load the XLSX workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Use default HTML save options (preserves all content)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Export the workbook to an HTML file
        workbook.Save("output.html", htmlOptions);
    }
}
