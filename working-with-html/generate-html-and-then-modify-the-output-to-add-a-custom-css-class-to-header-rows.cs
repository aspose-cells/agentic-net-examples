// Title: Export Excel to HTML with Aspose.Cells and Apply a Custom CSS Class to Header Rows (C#)
// Description: Creates a workbook, adds sample data, configures HtmlSaveOptions with a CSS rule, saves as HTML, then replaces the first <tr> tag with <tr class="my‑header"> to style the header row before writing the file back.
// Keywords: Aspose.Cells | C# | .NET | export Excel to HTML | HtmlSaveOptions | custom CSS class | header row styling | modify generated HTML | table row class injection | Excel web report
// Common Searches: Aspose.Cells add CSS class to header row | C# export Excel as HTML with custom styles | How to style header rows in Aspose.Cells HTML output | Replace <tr> tag in Aspose.Cells generated HTML | Add custom CSS to Aspose.Cells HTML export
// Developer Intent: Apply a custom CSS class to the header row of the HTML file produced by Aspose.Cells in a C# application.
// Use Cases: Generate web‑ready reports from Excel where the header row needs distinct visual emphasis. | Create HTML email templates from workbooks with highlighted column titles. | Batch‑convert multiple worksheets to HTML while enforcing a consistent header style across all pages. | Integrate Excel‑derived tables into existing web pages that use a shared CSS framework.
// AI Prompts: Write C# code that uses Aspose.Cells to export a workbook to HTML and inject a custom CSS class into the first table row. | Show how to read the HTML output from Aspose.Cells, locate header <tr> elements, and add a specified CSS class to each. | Explain configuring HtmlSaveOptions to embed custom CSS and then post‑process the generated HTML to apply a class to header rows.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, adds sample data, configures HtmlSaveOptions with a CSS rule, saves as HTML, then replaces the first <tr> tag with <tr class="my‑header"> to style the header row before writing the file back.
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

        // Configure HTML save options with a custom CSS rule for header rows
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.CssStyles = ".my-header { background-color:#e0e0e0; font-weight:bold; }";

        // Save the workbook as HTML
        string htmlFilePath = "output.html";
        workbook.Save(htmlFilePath, saveOptions);

        // Load the generated HTML content
        string htmlContent = File.ReadAllText(htmlFilePath);

        // Add a custom CSS class to the first <tr> element (assumed header row)
        int firstTrIndex = htmlContent.IndexOf("<tr>", StringComparison.Ordinal);
        if (firstTrIndex >= 0)
        {
            htmlContent = htmlContent.Remove(firstTrIndex, 4)
                                     .Insert(firstTrIndex, "<tr class=\"my-header\">");
        }

        // Write the modified HTML back to the file
        File.WriteAllText(htmlFilePath, htmlContent);

        Console.WriteLine("HTML file created with custom header class.");
    }
}
