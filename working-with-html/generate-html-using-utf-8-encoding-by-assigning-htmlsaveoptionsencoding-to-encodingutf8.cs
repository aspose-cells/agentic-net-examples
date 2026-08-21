// Title: Save Aspose.Cells Workbook as UTF-8 HTML using HtmlSaveOptions.Encoding (C#)
// Description: Demonstrates how to create a workbook, insert Unicode text (including emojis), set HtmlSaveOptions.Encoding to Encoding.UTF8, and export the workbook to a UTF‑8 encoded HTML file with Aspose.Cells for C#.
// Keywords: Aspose.Cells HTML export | UTF-8 encoding C# | HtmlSaveOptions.Encoding | Unicode Excel to HTML | emoji support Aspose.Cells | C# save workbook as HTML | global character set Aspose
// Common Searches: Aspose.Cells export to HTML with UTF-8 | Set HtmlSaveOptions.Encoding to UTF-8 in C# | How to preserve Unicode characters when saving Excel as HTML | C# Aspose.Cells HTML output encoding | Save workbook as UTF-8 HTML file
// Developer Intent: Generate an HTML file from an Aspose.Cells workbook that uses UTF‑8 encoding to correctly render Unicode characters.
// Use Cases: Publish multilingual Excel reports on the web with proper Unicode display. | Create HTML email templates from spreadsheets that include special symbols or emojis. | Build web‑based dashboards that require UTF‑8 encoded HTML for international audiences.
// AI Prompts: Provide C# code that sets HtmlSaveOptions.Encoding to UTF-8 when saving an Aspose.Cells workbook as HTML. | Show an example of exporting a workbook containing emojis to an HTML file with Aspose.Cells and verifying the UTF-8 charset. | Explain how to check the character encoding of the generated HTML file after using HtmlSaveOptions.Encoding = Encoding.UTF8.

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsHtmlEncodingDemo
{
    // Demonstrates how to create a workbook, insert Unicode text (including emojis), set HtmlSaveOptions.Encoding to Encoding.UTF8, and export the workbook to a UTF‑8 encoded HTML file with Aspose.Cells for C#.
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

            // Save the workbook as an HTML file using the specified encoding
            string outputPath = "output_utf8.html";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"HTML file saved to {outputPath} with UTF-8 encoding.");
        }
    }
}
