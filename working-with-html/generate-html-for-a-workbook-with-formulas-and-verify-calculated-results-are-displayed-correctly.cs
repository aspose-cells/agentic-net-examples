// Title: Generate HTML from a workbook with a SUM formula evaluated using Aspose.Cells for .NET
// AI Prompts: Write C# code that inserts numeric values, adds a SUM formula, forces formula calculation, and saves the active worksheet as an HTML file with Aspose.Cells. | Show how to set HtmlSaveOptions to export only the active sheet after formulas have been evaluated in a workbook. | Demonstrate checking that the calculated result of a formula appears correctly in the generated HTML output using Aspose.Cells.
// Common Searches: Aspose.Cells .NET export worksheet to HTML after calculating formulas | C# save Excel workbook as HTML with evaluated SUM formula using Aspose.Cells | HtmlSaveOptions ExportActiveWorksheetOnly example with formula calculation | display calculated cell values in HTML output from Aspose.Cells | generate HTML from workbook that contains formulas in C#
// Tags: calculate workbook formulas before HTML export | HtmlSaveOptions ExportActiveWorksheetOnly usage | Aspose.Cells SUM formula to HTML rendering | C# workbook to HTML with evaluated formulas | export active worksheet as HTML Aspose.Cells

using System;
using Aspose.Cells;

// The program creates a workbook, fills cells A1‑A3 with numbers, inserts a SUM formula in B1, calculates all formulas, and saves only the active worksheet as an HTML file using HtmlSaveOptions, ensuring the calculated result is displayed.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate cells with numeric values
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].PutValue(30);

        // Insert a formula that sums the three values
        sheet.Cells["B1"].Formula = "=SUM(A1:A3)";

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Configure HTML save options to export the active worksheet only
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportActiveWorksheetOnly = true
        };

        // Save the workbook as an HTML file; the calculated result will be displayed in the HTML
        workbook.Save("output.html", htmlOptions);
    }
}
