// Title: Export an Excel workbook to UTF-8 encoded HTML with Aspose.Cells for .NET (C#)
// Common Searches: how to export Excel to UTF-8 HTML using Aspose.Cells in C# | Aspose.Cells C# set HtmlSaveOptions.Encoding to UTF8 | save workbook as HTML with UTF-8 encoding Aspose.Cells .NET | generate UTF-8 HTML file from Excel workbook using Aspose.Cells API | C# code sample for saving workbook as UTF-8 HTML with Aspose.Cells
// Tags: HtmlSaveOptions UTF-8 configuration Aspose.Cells | Excel to HTML conversion .NET Aspose.Cells | specify output encoding for HTML export Aspose.Cells | workbook HTML export with UTF-8 charset | C# Aspose.Cells HTML save options example

using System.Text;
using Aspose.Cells;

// The example creates a Workbook, writes values to cells, configures HtmlSaveOptions.Encoding to UTF-8, and saves the workbook as an HTML file named output.html.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Example: add some data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Configure HTML save options with UTF-8 encoding
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.Encoding = Encoding.UTF8; // Set UTF-8 encoding

        // Save the workbook as an HTML file using the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
