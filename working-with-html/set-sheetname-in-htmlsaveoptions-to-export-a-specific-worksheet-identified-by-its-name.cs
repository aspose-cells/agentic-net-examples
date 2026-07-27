// Title: Export a Single Worksheet to HTML with Aspose.Cells – Set SheetSet in HtmlSaveOptions (C#)
// Description: Learn how to render only a chosen worksheet to HTML using Aspose.Cells for .NET. The example creates a workbook, assigns a name to the target sheet, configures HtmlSaveOptions.SheetSet with that name, and saves the result as an HTML file containing just that sheet.
// Keywords: Aspose.Cells export single sheet HTML | HtmlSaveOptions SheetSet C# | save specific worksheet as HTML | Aspose.Cells HTML conversion by sheet name | .NET render one worksheet to HTML
// Common Searches: Aspose.Cells export only one worksheet to HTML | HtmlSaveOptions SheetSet property example | C# save selected sheet as HTML using Aspose.Cells | How to render a specific worksheet to HTML with Aspose
// Developer Intent: Render only the worksheet named "Sheet2" to an HTML file.
// Use Cases: Display a single report sheet on a web page without loading the entire workbook. | Generate separate HTML files for each tab in a large Excel file to improve page‑load performance. | Provide a downloadable HTML version of a confidential analysis sheet while keeping other data hidden.
// AI Prompts: Show how to export multiple worksheets to HTML by passing a list of sheet names to HtmlSaveOptions.SheetSet. | Write code that reads a sheet name from user input and dynamically sets HtmlSaveOptions.SheetSet to export that sheet. | Create a loop that saves every worksheet in a workbook as an individual HTML file using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Learn how to render only a chosen worksheet to HTML using Aspose.Cells for .NET. The example creates a workbook, assigns a name to the target sheet, configures HtmlSaveOptions.SheetSet with that name, and saves the result as an HTML file containing just that sheet.
class ExportSpecificWorksheetToHtml
{
    static void Main()
    {
        // Create a new workbook and add two worksheets
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Name = "Sheet1";
        workbook.Worksheets[0].Cells["A1"].PutValue("Data in Sheet1");

        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
        sheet2.Cells["A1"].PutValue("Data in Sheet2");

        // Initialize HtmlSaveOptions
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Export only the worksheet named "Sheet2" by setting SheetSet with the sheet name
        htmlOptions.SheetSet = new SheetSet("Sheet2");

        // Save the workbook to HTML; only "Sheet2" will be rendered
        workbook.Save("Sheet2.html", htmlOptions);
    }
}
