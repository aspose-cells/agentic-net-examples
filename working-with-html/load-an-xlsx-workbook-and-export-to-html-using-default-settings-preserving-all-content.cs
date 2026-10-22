// Title: Export an XLSX workbook to HTML using Aspose.Cells in C# with default HtmlSaveOptions
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells and saves it as .html using the default HtmlSaveOptions. | Show how to convert an Excel workbook to an HTML page in .NET without customizing any save options. | Provide a minimal C# snippet that reads a spreadsheet and writes it to HTML while preserving all content.
// Common Searches: asp.net how to export an xlsx workbook to html using aspose.cells default settings | c# Aspose.Cells save workbook as html preserving formulas and images | convert Excel file to web page with Aspose.Cells without specifying HtmlSaveOptions | default HTML export options in Aspose.Cells for .NET | sample code to load input.xlsx and generate output.html with Aspose.Cells
// Tags: Aspose.Cells HTML export default options | C# convert XLSX to HTML Aspose.Cells | preserve Excel content when saving as HTML .NET | Workbook.Save HTML example Aspose.Cells | export workbook to web page using Aspose.Cells

using System;
using Aspose.Cells;

// Loads 'input.xlsx' into a Workbook, creates a default HtmlSaveOptions object, and saves the workbook as 'output.html', preserving all worksheet content.
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
