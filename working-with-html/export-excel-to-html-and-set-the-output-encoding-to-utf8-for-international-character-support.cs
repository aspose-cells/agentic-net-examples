// Title: Convert an Excel workbook to UTF‑8 encoded HTML using Aspose.Cells for .NET (C#)
// Common Searches: c# asp.net convert xlsx file to html with utf-8 encoding using aspose.cells | how to preserve Unicode characters when exporting Excel to HTML with Aspose.Cells | set HtmlSaveOptions.Encoding to UTF8 for Excel to HTML conversion in C# | asp.net core save workbook as html ensuring international characters are retained
// Tags: Aspose.Cells HtmlSaveOptions UTF‑8 | Excel to HTML conversion C# | Unicode support Aspose.Cells HTML export | HTML output configuration Aspose.Cells | Save workbook as HTML with international characters

using System;
using System.Text;
using Aspose.Cells;

// The example loads an Excel workbook, configures HtmlSaveOptions to use UTF‑8 encoding, and saves the file as HTML, guaranteeing that all international (Unicode) characters are correctly rendered in the output.
class Program
{
    static void Main()
    {
        // Load the source Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to use UTF‑8 encoding
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.Encoding = Encoding.UTF8; // ensures international characters are preserved

        // Save the workbook as an HTML file with the specified encoding
        workbook.Save("output.html", htmlOptions);
    }
}
