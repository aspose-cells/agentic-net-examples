// Title: Export Excel to HTML with Calculated Values (no Formula Text) using Aspose.Cells for .NET
// Description: Shows how to save a workbook as HTML with Aspose.Cells, evaluating formulas (CalculateFormula = true) and suppressing formula strings (ExportFormula = false) so the output displays only the computed results.
// Keywords: Aspose.Cells | HTML export | CalculateFormula | ExportFormula | C# Excel to HTML | hide formulas | formula results | static HTML report | Aspose.Cells HtmlSaveOptions
// Common Searches: Aspose.Cells export HTML calculated values | C# save Excel as HTML without formulas | HtmlSaveOptions ExportFormula false | show formula results in HTML Aspose | convert Excel to HTML values only
// Developer Intent: Create an HTML file from an Excel workbook where formulas are evaluated and only the resulting values are shown.
// Use Cases: Produce web‑ready reports that display computed numbers without exposing underlying formulas. | Generate static HTML snapshots of spreadsheets for newsletters or documentation. | Publish dashboard pages where Excel calculations appear as plain values for end users.
// AI Prompts: Write C# code to export each worksheet of a workbook to separate HTML files while keeping calculated values. | Demonstrate how to retain cell styles and number formats in the HTML export while hiding formulas. | Explain how to toggle ExportFormula at runtime based on a user‑selected option in an Aspose.Cells HTML export workflow.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Shows how to save a workbook as HTML with Aspose.Cells, evaluating formulas (CalculateFormula = true) and suppressing formula strings (ExportFormula = false) so the output displays only the computed results.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some data with formulas
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["B1"].PutValue(20);
            sheet.Cells["C1"].Formula = "=A1+B1"; // Formula to be displayed as value

            // Configure HTML save options:
            // - CalculateFormula = true ensures formulas are evaluated before saving.
            // - ExportFormula = false prevents the formula text from being written to HTML.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                CalculateFormula = true,
                ExportFormula = false
            };

            // Save the workbook as HTML with the specified options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Excel exported to HTML with formulas shown as values.");
        }
    }
}
