// Title: Aspose.Cells for .NET – Set Times New Roman as default HTML font
// Description: Learn how to use Aspose.Cells HtmlSaveOptions in C# to set the DefaultFontName property to "Times New Roman" so that the generated HTML file renders all worksheet text with that font. The example creates a workbook, writes sample text, configures the option, and saves the file.
// Keywords: Aspose.Cells C# | HtmlSaveOptions DefaultFontName | set default HTML font Aspose.Cells | Times New Roman HTML export | Aspose.Cells .NET HTML rendering | save workbook as HTML font | C# Aspose.Cells example
// Common Searches: Aspose.Cells set default font for HTML export | HtmlSaveOptions DefaultFontName C# example | How to force Times New Roman in Aspose.Cells HTML output | Export Excel to HTML with specific font using Aspose.Cells | C# Aspose.Cells HTMLSaveOptions font settings
// Developer Intent: Specify Times New Roman as the default font for HTML rendering when saving a workbook with Aspose.Cells.
// Use Cases: Corporate reporting that must follow a Times New Roman style guide. | Web‑based spreadsheet viewers where a consistent font improves readability. | Email‑ready HTML spreadsheets that match branding without extra CSS.
// AI Prompts: Show C# code that sets HtmlSaveOptions.DefaultFontName to Times New Roman in Aspose.Cells. | Explain how Aspose.Cells falls back to a substitute font if Times New Roman is unavailable on the server. | Provide an example that sets the default font and adds a custom CSS file to the HTML export.

using System;
using Aspose.Cells;

// Learn how to use Aspose.Cells HtmlSaveOptions in C# to set the DefaultFontName property to "Times New Roman" so that the generated HTML file renders all worksheet text with that font. The example creates a workbook, writes sample text, configures the option, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample data to demonstrate the default font
        worksheet.Cells["A1"].PutValue("This text will be rendered with Times New Roman as the default font.");

        // Configure HTML save options and set the default font name
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.DefaultFontName = "Times New Roman";

        // Save the workbook to an HTML file using the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
