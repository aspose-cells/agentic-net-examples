// Title: Export Excel to HTML with a Custom <title> Using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create or load an Excel workbook, set its BuiltInDocumentProperties.Title, configure HtmlSaveOptions.PageTitle, and save the file as HTML so the generated <title> element reflects the workbook title.
// Keywords: Aspose.Cells | HTML export | custom page title | HtmlSaveOptions.PageTitle | BuiltInDocumentProperties.Title | C# Excel to HTML | Aspose.Cells .NET | Excel workbook HTML header | Aspose.Cells export example
// Common Searches: Aspose.Cells set HTML title | C# export Excel to HTML with custom title | HtmlSaveOptions PageTitle property C# | How to add <title> tag when saving Excel as HTML Aspose | Use BuiltInDocumentProperties.Title for HTML output
// Developer Intent: I need to convert an Excel workbook to an HTML file and specify a custom <title> element in the HTML header.
// Use Cases: Embedding Excel data in web pages with SEO‑friendly titles | Generating printable reports where the browser tab shows the workbook name | Batch converting multiple spreadsheets to HTML, each inheriting its own title property | Creating downloadable HTML versions of financial models with clear page titles
// AI Prompts: Write C# code that loads an existing .xlsx, sets BuiltInDocumentProperties.Title to a given string, and saves as HTML with that title using Aspose.Cells. | Show how to add meta description and charset tags alongside PageTitle when exporting Excel to HTML with Aspose.Cells. | Provide a script to process all .xlsx files in a folder, assign each file name as the HTML <title>, and save the HTML files using Aspose.Cells. | Explain the difference between HtmlSaveOptions.PageTitle and Workbook.BuiltInDocumentProperties.Title when exporting to HTML.

using System;
using Aspose.Cells;

// Demonstrates how to create or load an Excel workbook, set its BuiltInDocumentProperties.Title, configure HtmlSaveOptions.PageTitle, and save the file as HTML so the generated <title> element reflects the workbook title.
class ExportExcelToHtmlWithTitle
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet and add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "SampleSheet";
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Set the workbook title property (this will be used as the HTML page title)
        workbook.BuiltInDocumentProperties.Title = "My Custom Workbook Title";

        // Configure HTML save options and assign the custom page title
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.PageTitle = workbook.BuiltInDocumentProperties.Title; // custom HTML header

        // Save the workbook as an HTML file with the specified title
        workbook.Save("output.html", saveOptions);
    }
}
