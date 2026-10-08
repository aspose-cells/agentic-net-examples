// Title: Remove trailing whitespace after <br> tags in HTML generated from an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, saves it as HTML, and then uses a regular expression to strip any whitespace characters that appear immediately after <br> elements. | Provide a Regex pattern and C# implementation to collapse spaces, tabs, or line breaks following <br> tags in an HTML string produced by Aspose.Cells. | Show how to read the HTML output from a MemoryStream, apply post‑processing to clean up <br> tag spacing, and write the sanitized HTML to a file.
// Common Searches: aspocells c# remove spaces after <br> in saved html | how to clean up extra whitespace after line break tags when exporting Excel to HTML | regex to trim whitespace after <br> in Aspose.Cells HTML output | post processing Aspose.Cells HTML to eliminate blank characters after <br> | c# Aspose.Cells HtmlSaveOptions remove trailing spaces after br tag
// Tags: Aspose.Cells HTML post‑processing | C# regex whitespace after br tag | remove trailing spaces in Aspose.Cells HTML output | HtmlSaveOptions cleanup of line break spacing | Excel to HTML conversion whitespace handling

using Aspose.Cells;
using System;
using System.IO;
using System.Text.RegularExpressions;

// The example loads an Excel workbook with Aspose.Cells, saves it to a MemoryStream as HTML using HtmlSaveOptions, reads the HTML into a string, applies a case‑insensitive regular expression to replace any whitespace characters that follow <br> tags with a plain <br>, and writes the cleaned HTML to an output file.
class Program
{
    static void Main()
    {
        // Load the workbook from an existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options as needed
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        // Example: export only the active worksheet
        // htmlOptions.ExportActiveWorksheetOnly = true;

        // Save the workbook to a memory stream in HTML format
        using (MemoryStream htmlStream = new MemoryStream())
        {
            workbook.Save(htmlStream, htmlOptions);
            htmlStream.Position = 0;

            // Read the generated HTML into a string
            string htmlContent = new StreamReader(htmlStream).ReadToEnd();

            // Remove unnecessary blank spaces after <br> tags
            // This replaces any whitespace characters following a <br> with nothing
            htmlContent = Regex.Replace(htmlContent, @"<br>\s+", "<br>", RegexOptions.IgnoreCase);

            // Write the cleaned HTML to an output file
            File.WriteAllText("output.html", htmlContent);
        }
    }
}
