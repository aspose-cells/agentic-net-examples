// Title: Export Excel to HTML with a Custom <title> Tag Using Aspose.Cells for .NET
// Description: This C# sample builds a workbook, populates cells, assigns a value to the workbook's Title property, sets HtmlSaveOptions.PageTitle, and saves the file as HTML so the generated <title> element mirrors the workbook title.
// Keywords: Aspose.Cells | HTML export | PageTitle | HtmlSaveOptions | .NET | Excel to HTML | custom title tag | BuiltInDocumentProperties | C# example
// Common Searches: Aspose.Cells set HTML title tag | HtmlSaveOptions PageTitle C# | Export Excel as HTML with custom header Aspose | How to add <title> to HTML output from workbook | C# Aspose.Cells HTML export custom page title
// Developer Intent: Add a specific <title> element to the HTML file generated from an Excel workbook.
// Use Cases: Generate web‑ready reports where the browser tab displays the workbook name. | Batch‑export several spreadsheets, each using its own Title property for SEO‑friendly page headings. | Integrate the exported HTML into a CMS that relies on the <title> tag for navigation and indexing.
// AI Prompts: Generate C# code that uses Aspose.Cells to save a workbook as HTML and sets the page title from BuiltInDocumentProperties.Title. | Show how to customize the HTML head section (title, meta description, charset) when exporting Excel with Aspose.Cells. | Explain how HtmlSaveOptions can be combined with other options like ExportImagesAsBase64 for a complete HTML report.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // This C# sample builds a workbook, populates cells, assigns a value to the workbook's Title property, sets HtmlSaveOptions.PageTitle, and saves the file as HTML so the generated <title> element mirrors the workbook title.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John Doe");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Jane Smith");
            sheet.Cells["B3"].PutValue(28);

            // Set the workbook title property (optional, can be any custom string)
            string customTitle = "My Custom Workbook Title";
            workbook.BuiltInDocumentProperties.Title = customTitle;

            // Configure HTML save options and assign the page title
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            // This property sets the <title> element in the generated HTML file
            saveOptions.PageTitle = customTitle;

            // Save the workbook as an HTML file with the custom page title
            string outputPath = "ExportedWorkbook.html";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"HTML file saved to '{outputPath}' with page title '{customTitle}'.");
        }
    }
}
