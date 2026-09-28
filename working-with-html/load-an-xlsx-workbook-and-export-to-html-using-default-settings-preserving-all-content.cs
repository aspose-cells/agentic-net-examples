// Title: Export an XLSX workbook to HTML with Aspose.Cells for .NET using default settings
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells and saves it as an .html file using the library's default export options. | Show how to convert an Excel workbook to a fully formatted HTML page in a .NET application with Aspose.Cells without customizing any save parameters.
// Common Searches: aspnet convert xlsx file to html using Aspose.Cells default configuration | c# how to save Excel workbook as html preserving styles with Aspose.Cells | Aspose.Cells export workbook to html without specifying HtmlSaveOptions | default HTML output from Aspose.Cells for .NET example | save spreadsheet as web page using Aspose.Cells C#
// Tags: Aspose.Cells workbook.Save to Html | C# XLSX to HTML conversion Aspose.Cells | default HTML export Aspose.Cells .NET | preserve Excel formatting HTML Aspose.Cells | Aspose.Cells SaveFormat.Html usage

using Aspose.Cells;

// The sample loads "input.xlsx" into an Aspose.Cells Workbook object and calls workbook.Save with SaveFormat.Html, producing "output.html" using the library's default HTML export settings while retaining all workbook content and formatting.
class Program
{
    static void Main()
    {
        // Load the XLSX workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Export the workbook to HTML using default settings (preserves all content)
        workbook.Save("output.html", SaveFormat.Html);
    }
}
