// Title: Export Excel to HTML with Row/Column Heading Anchors and Worksheet Hyperlink Navigation – Aspose.Cells for .NET
// Description: This C# example builds a workbook with an index sheet and a Details sheet, adds a hyperlink from the index to the Details cell, and saves the file as HTML using HtmlSaveOptions (ExportRowColumnHeadings = true, ExportNamedRangeAnchors = true). The output contains anchor tags for row and column headings, enabling click‑through navigation between sections.
// Keywords: Aspose.Cells | HTML export | ExportRowColumnHeadings | ExportNamedRangeAnchors | worksheet hyperlink | C# Excel to HTML | anchor navigation | row heading anchors | column heading anchors | cross‑sheet link
// Common Searches: Aspose.Cells export workbook to HTML with navigation | How to add hyperlink between worksheets in HTML output | Enable row and column heading anchors in Aspose.Cells HTML | HtmlSaveOptions ExportNamedRangeAnchors example | C# create clickable index in exported HTML Excel
// Developer Intent: Create an HTML representation of an Excel file where index entries link to detailed sections via anchors generated from row/column headings and worksheet hyperlinks.
// Use Cases: Generate web‑ready reports that let users jump from a summary table to detailed pages. | Automate documentation where an index sheet provides quick navigation to content sheets. | Validate that exported HTML preserves Excel hyperlink behavior for QA testing.
// AI Prompts: Write a C# unit test that loads the saved HTML and verifies the hyperlink target ID matches the Details sheet anchor. | Describe the HTML markup produced when ExportRowColumnHeadings and ExportNamedRangeAnchors are enabled. | Show how to rename the generated anchor IDs for rows and columns using HtmlSaveOptions callbacks.

using System;
using System.IO;
using Aspose.Cells;

// This C# example builds a workbook with an index sheet and a Details sheet, adds a hyperlink from the index to the Details cell, and saves the file as HTML using HtmlSaveOptions (ExportRowColumnHeadings = true, ExportNamedRangeAnchors = true). The output contains anchor tags for row and column headings, enabling click‑through navigation between sections.
class HtmlNavigationDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet mainSheet = workbook.Worksheets[0];
            mainSheet.Name = "Main";

            // Add row and column headings
            mainSheet.Cells["A1"].PutValue("Index");
            mainSheet.Cells["B1"].PutValue("Section");
            mainSheet.Cells["A2"].PutValue("1");
            mainSheet.Cells["B2"].PutValue("Introduction");
            mainSheet.Cells["A3"].PutValue("2");
            mainSheet.Cells["B3"].PutValue("Details");

            // Add a second worksheet that will be the target of a hyperlink
            Worksheet detailsSheet = workbook.Worksheets.Add("Details");
            detailsSheet.Cells["A1"].PutValue("Details Section");
            detailsSheet.Cells["A2"].PutValue("More information here...");

            // Create a hyperlink in the main sheet that points to the Details sheet
            // Parameters: firstCell, totalRows, totalColumns, address
            mainSheet.Hyperlinks.Add("B2", 1, 1, "Details!A1");

            // Configure HTML save options:
            // - ExportRowColumnHeadings: exports row/column headings as anchors, enabling navigation
            // - ExportNamedRangeAnchors: keeps anchor elements for named ranges (optional but safe)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportRowColumnHeadings = true,
                ExportNamedRangeAnchors = true
            };

            // Define output file path
            string outputPath = "NavigationDemo.html";

            // Save the workbook as an HTML file
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
