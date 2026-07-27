// Title: Export All Worksheets to One HTML File with Aspose.Cells for .NET
// Description: Shows how to configure Aspose.Cells HtmlSaveOptions (SaveAsSingleFile = true, ShowAllSheets = true) to convert a multi‑sheet workbook into a single HTML document.
// Keywords: Aspose.Cells | C# | .NET | HtmlSaveOptions | SaveAsSingleFile | ShowAllSheets | export all worksheets to HTML | single HTML output | multi‑sheet HTML export | spreadsheet to HTML conversion
// Common Searches: Aspose.Cells export all sheets to single HTML | C# HtmlSaveOptions ShowAllSheets example | Save workbook as one HTML page Aspose | How to generate single HTML file from multiple worksheets .NET | Aspose.Cells SaveAsSingleFile true
// Developer Intent: Generate one HTML file that includes every worksheet from a workbook.
// Use Cases: Create a web‑ready report that consolidates data from several sheets into a single page. | Provide a printable HTML version of a complete workbook for distribution. | Embed a full workbook preview in a web application without loading multiple files.
// AI Prompts: How can I add a custom CSS stylesheet when exporting all worksheets to a single HTML file with Aspose.Cells? | Show an example of exporting only selected worksheets to one HTML document using HtmlSaveOptions. | Explain how to embed cell images while saving the entire workbook as a single HTML page. | What options are available to control table layout and fonts when generating a single HTML file from multiple sheets?

using System;
using Aspose.Cells;

// Shows how to configure Aspose.Cells HtmlSaveOptions (SaveAsSingleFile = true, ShowAllSheets = true) to convert a multi‑sheet workbook into a single HTML document.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add data to the first worksheet
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "FirstSheet";
        sheet1.Cells["A1"].PutValue("Content of Sheet 1");

        // Add a second worksheet
        Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
        sheet2.Cells["A1"].PutValue("Content of Sheet 2");

        // Configure HTML save options to generate a single file with all sheets
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.SaveAsSingleFile = true;   // Save as a single HTML file
        saveOptions.ShowAllSheets = true;      // Include all worksheets in the output

        // Save the workbook as HTML
        workbook.Save("AllSheets.html", saveOptions);
    }
}
