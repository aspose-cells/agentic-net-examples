// Title: C# AspNet – Export Excel to HTML with Clickable Row/Column Headings and Named‑Range Anchors using Aspose.Cells
// Description: Shows how to build a workbook, add row and column titles, assign named ranges, and save it as HTML with HtmlSaveOptions configured to turn those titles into intra‑page navigation anchors.
// Keywords: Aspose.Cells | C# | HTML export | ExportRowColumnHeadings | ExportNamedRangeAnchors | named range anchors | clickable headings | Excel to HTML | navigation links | web‑friendly spreadsheet
// Common Searches: Aspose.Cells export headings as links | HTML output with named range anchors Aspose | C# generate clickable row column headings in HTML | how to add navigation anchors when saving Excel as HTML | Aspose.Cells HTMLSaveOptions navigation example
// Developer Intent: Create an HTML file from an Excel workbook where row and column titles serve as internal navigation links.
// Use Cases: Build a report workbook, mark column titles with named ranges, and export to HTML so users can jump to sections by clicking headings. | Validate that the resulting HTML contains <a name="Heading1"> and corresponding <a href="#Heading1"> elements for quick navigation. | Integrate the export step into an automated pipeline that delivers web‑ready spreadsheets with built‑in navigation.
// AI Prompts: Generate C# code with Aspose.Cells that exports a worksheet to HTML and makes row/column headings clickable. | Explain the effect of ExportRowColumnHeadings and ExportNamedRangeAnchors on the HTML markup produced by Aspose.Cells. | Write a unit test in C# that loads the exported HTML and asserts the presence of navigation anchors for the defined headings.

using System;
using Aspose.Cells;

// Shows how to build a workbook, add row and column titles, assign named ranges, and save it as HTML with HtmlSaveOptions configured to turn those titles into intra‑page navigation anchors.
class TestNavigationBetweenHeadings
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Sheet1";

        // Add column headings (these will become navigation anchors)
        sheet.Cells["B1"].PutValue("Heading 1");
        sheet.Cells["C1"].PutValue("Heading 2");

        // Add row headings (these will also be exported as navigation links)
        sheet.Cells["A2"].PutValue("Row 1");
        sheet.Cells["A3"].PutValue("Row 2");

        // Fill some sample data
        sheet.Cells["B2"].PutValue("Data 1");
        sheet.Cells["C2"].PutValue("Data 2");
        sheet.Cells["B3"].PutValue("Data 3");
        sheet.Cells["C3"].PutValue("Data 4");

        // Create named ranges for the column headings – they will be exported as <a> anchors
        sheet.Cells.CreateRange("B1", "B1").Name = "Heading1";
        sheet.Cells.CreateRange("C1", "C1").Name = "Heading2";

        // Configure HTML save options to export row/column headings and named‑range anchors
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportRowColumnHeadings = true;   // export row/column headings as clickable links
        htmlOptions.ExportNamedRangeAnchors = true;   // generate <a name="..."> anchors for named ranges

        // Save the workbook as an HTML file
        workbook.Save("NavigationHeadings.html", htmlOptions);
    }
}
