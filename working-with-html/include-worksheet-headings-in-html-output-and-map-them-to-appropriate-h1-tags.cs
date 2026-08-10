// Title: Convert Excel Header Cells to <h1> Tags in HTML Export with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds header and data rows, enables ExportRowColumnHeadings, saves the sheet as HTML, then uses a C# regular expression to replace all <th> elements with <h1> tags and writes the result to a new file. Ideal for generating SEO‑friendly HTML from Excel.
// Keywords: Aspose.Cells | C# | HTML export | ExportRowColumnHeadings | Excel to HTML | convert th to h1 | worksheet headings | SEO friendly HTML | HtmlSaveOptions | regex replace
// Common Searches: Aspose.Cells export Excel to HTML with row and column headings | C# replace <th> with <h1> in Aspose.Cells generated HTML | How to include worksheet headings in Aspose.Cells HTML output | Map Excel column headers to H1 tags using Aspose.Cells | Convert Excel table headers to H1 in HTML C#
// Developer Intent: Generate an HTML file from an Excel workbook where the worksheet’s header cells are rendered as <h1> elements instead of table header cells.
// Use Cases: Create SEO‑optimized web reports from Excel data by turning header cells into top‑level headings. | Improve accessibility by presenting section titles as <h1> tags for screen readers. | Integrate Excel‑derived content into web applications that require heading tags for styling or navigation. | Produce printable HTML pages where row/column labels serve as document titles.
// AI Prompts: Show C# code that exports an Excel worksheet to HTML with row/column headings and replaces all <th> tags with <h1> tags. | Explain how to configure Aspose.Cells HtmlSaveOptions to include headings and then transform them using a regular expression. | Provide a step‑by‑step guide to create SEO‑friendly HTML from Excel using Aspose.Cells for .NET.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// Creates a workbook, adds header and data rows, enables ExportRowColumnHeadings, saves the sheet as HTML, then uses a C# regular expression to replace all <th> elements with <h1> tags and writes the result to a new file. Ideal for generating SEO‑friendly HTML from Excel.
class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Header1");
        worksheet.Cells["B1"].PutValue("Header2");
        worksheet.Cells["A2"].PutValue("Data1");
        worksheet.Cells["B2"].PutValue("Data2");

        // Configure HTML save options to include row and column headings
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportRowColumnHeadings = true; // Export A, B, 1, 2 headings

        // Save the workbook as HTML
        string htmlFilePath = "output.html";
        workbook.Save(htmlFilePath, saveOptions);

        // Read the generated HTML
        string htmlContent = File.ReadAllText(htmlFilePath);

        // Replace heading cells (<th> elements) with <h1> tags
        string updatedHtml = Regex.Replace(
            htmlContent,
            @"<th[^>]*>(.*?)</th>",
            "<h1>$1</h1>",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);

        // Save the modified HTML
        string finalHtmlPath = "output_with_h1_headings.html";
        File.WriteAllText(finalHtmlPath, updatedHtml);

        Console.WriteLine($"HTML file with <h1> headings saved to: {finalHtmlPath}");
    }
}
