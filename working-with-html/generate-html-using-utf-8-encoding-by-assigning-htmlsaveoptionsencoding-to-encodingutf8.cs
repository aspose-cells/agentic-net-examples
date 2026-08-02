// Title: Save Aspose.Cells Workbook as UTF‑8 HTML in C#
// Description: Creates a new Workbook, writes Unicode text with an emoji to cell A1, sets HtmlSaveOptions.Encoding to Encoding.UTF8, and saves the result as output_utf8.html.
// Keywords: Aspose.Cells | C# | HtmlSaveOptions | UTF-8 | HTML export | Unicode | emoji | Encoding.UTF8 | Excel to HTML | save workbook as HTML
// Common Searches: Aspose.Cells save as UTF-8 HTML C# | HtmlSaveOptions Encoding UTF8 example | export Excel to HTML with Unicode characters | C# generate HTML from workbook with UTF-8 encoding | how to preserve emojis when saving Excel as HTML
// Developer Intent: Generate an HTML file from a workbook using UTF‑8 encoding with Aspose.Cells.
// Use Cases: Publish multilingual reports or dashboards on the web. | Create HTML email templates that include emojis or special characters. | Convert Excel‑based documentation to web‑ready pages without losing Unicode data.
// AI Prompts: Show how to configure HtmlSaveOptions.Encoding to UTF-8 when saving a workbook as HTML with Aspose.Cells in C#. | Provide C# code that exports an Aspose.Cells workbook to HTML, sets UTF-8 encoding, and attaches a custom stylesheet. | Explain methods to verify that the generated HTML file is correctly encoded in UTF-8.

using System;
using System.Text;
using Aspose.Cells;

// Creates a new Workbook, writes Unicode text with an emoji to cell A1, sets HtmlSaveOptions.Encoding to Encoding.UTF8, and saves the result as output_utf8.html.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello, UTF-8 🌍");

        // Initialize HTML save options and set the encoding to UTF-8
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.Encoding = Encoding.UTF8;

        // Save the workbook as an HTML file using the specified options
        string outputPath = "output_utf8.html";
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine($"HTML file saved with UTF-8 encoding to: {outputPath}");
    }
}
