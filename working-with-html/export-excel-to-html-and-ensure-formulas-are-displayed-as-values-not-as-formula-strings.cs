// Title: Export Excel to HTML with calculated values (no formulas) using Aspose.Cells for .NET
// Description: Shows how to build a workbook, insert data and a formula, then save it as HTML with Aspose.Cells' HtmlSaveOptions so that formulas are evaluated (CalculateFormula = true) and the generated HTML contains only the computed values (ExportFormula = false).
// Keywords: Aspose.Cells | C# HtmlSaveOptions | ExportFormula false | CalculateFormula true | Excel to HTML conversion | display formula results | static HTML from Excel | Aspose.Cells .NET | HTML export options | formula evaluation
// Common Searches: Aspose.Cells export Excel to HTML calculated values | C# HtmlSaveOptions CalculateFormula true | hide formulas in HTML output Aspose.Cells | export workbook as static HTML using Aspose | save Excel as HTML with values not formulas
// Developer Intent: Generate an HTML file from an Excel workbook where all formulas are pre‑calculated and only the resulting numbers are displayed.
// Use Cases: Publish web‑ready reports from Excel sheets with formulas already resolved, ensuring the HTML shows static numbers. | Create snapshot versions of financial models for distribution without revealing the underlying calculation logic. | Automate conversion of Excel worksheets into email‑friendly HTML newsletters while preserving only the computed results.
// AI Prompts: Provide C# code using Aspose.Cells to export a workbook to HTML with formulas evaluated and displayed as values. | Explain the impact of HtmlSaveOptions.CalculateFormula and ExportFormula on the HTML output in Aspose.Cells. | Show how to add custom CSS to the generated HTML while keeping formulas exported as static values.

using System;
using Aspose.Cells;

// Shows how to build a workbook, insert data and a formula, then save it as HTML with Aspose.Cells' HtmlSaveOptions so that formulas are evaluated (CalculateFormula = true) and the generated HTML contains only the computed values (ExportFormula = false).
class ExportExcelToHtml
{
    static void Main()
    {
        // Create a new workbook and add sample data with a formula
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["B1"].PutValue(20);
        sheet.Cells["C1"].Formula = "=A1+B1";

        // Configure HTML save options:
        // - CalculateFormula = true ensures formulas are evaluated before saving.
        // - ExportFormula = false makes the HTML contain the calculated values, not the formula strings.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            CalculateFormula = true,
            ExportFormula = false
        };

        // Save the workbook as an HTML file with the specified options.
        workbook.Save("output.html", htmlOptions);
    }
}
