// Title: Export a Single Worksheet to HTML with Aspose.Cells for .NET using HtmlSaveOptions.SheetSet
// Description: Shows how to save only the worksheet named "Details" from a workbook to an HTML file by setting HtmlSaveOptions.SheetSet in C#.
// Keywords: Aspose.Cells | C# | HtmlSaveOptions | SheetSet | export single worksheet to HTML | save specific sheet as HTML | Aspose.Cells HTML export | selective worksheet export | Aspose.Cells .NET example
// Common Searches: Aspose.Cells export one sheet to HTML | HtmlSaveOptions SheetSet C# example | save specific worksheet as HTML Aspose | how to export only selected worksheet using Aspose.Cells | Aspose.Cells HTML export by sheet name
// Developer Intent: Save only the "Details" worksheet from a workbook as an HTML file.
// Use Cases: Create a web‑ready report that includes just the detailed data sheet from a larger workbook. | Generate separate HTML files for each worksheet by looping through sheet names and applying SheetSet. | Allow end‑users to download individual sections of a spreadsheet as HTML in a web application.
// AI Prompts: Provide C# code that exports a specific worksheet to HTML using Aspose.Cells HtmlSaveOptions.SheetSet. | Show how to export multiple selected worksheets to separate HTML files with Aspose.Cells. | Explain the role of HtmlSaveOptions.SheetSet when exporting worksheets to HTML in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Shows how to save only the worksheet named "Details" from a workbook to an HTML file by setting HtmlSaveOptions.SheetSet in C#.
class ExportSpecificWorksheetToHtml
{
    static void Main()
    {
        // Load or create a workbook
        Workbook workbook = new Workbook(); // creates a new workbook

        // Add worksheets and sample data
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Summary";
        sheet1.Cells["A1"].PutValue("Summary data");

        Worksheet sheet2 = workbook.Worksheets.Add("Details");
        sheet2.Cells["A1"].PutValue("Details data");

        // Prepare HtmlSaveOptions and specify the sheet to export by name
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        // Use SheetSet with the desired sheet name
        htmlOptions.SheetSet = new SheetSet("Details");

        // Save only the "Details" worksheet to HTML
        workbook.Save("DetailsOnly.html", htmlOptions);
    }
}
