// Title: Export Workbook with Formulas to HTML and Verify Calculated Result (Aspose.Cells C#)
// Description: Demonstrates how to create a workbook, add numeric data and a SUM formula, configure HtmlSaveOptions to calculate formulas and omit the original expression, save as HTML, and programmatically confirm that the computed value (30) appears in the generated markup.
// Keywords: Aspose.Cells HTML export | CalculateFormula option | ExportFormula false | C# verify HTML formula result | save Excel as static HTML | formula evaluation Aspose.Cells
// Common Searches: Aspose.Cells export Excel to HTML with calculated formulas | How to show formula results instead of formulas in HTML output | C# HtmlSaveOptions CalculateFormula example | Validate that HTML contains computed cell value | Generate static HTML report from workbook using Aspose
// Developer Intent: Generate an HTML file from a workbook where formulas are evaluated and only the resulting values are displayed, then programmatically ensure the expected result is present.
// Use Cases: Create web‑ready financial dashboards that display pre‑calculated totals as plain numbers. | Automated testing of spreadsheet‑to‑HTML conversion to guarantee formula accuracy. | Produce email‑friendly static HTML snapshots of Excel reports without exposing underlying formulas.
// AI Prompts: Write C# code with Aspose.Cells that saves a workbook containing formulas to HTML, enabling CalculateFormula and disabling ExportFormula. | Show how to read the saved HTML file and assert that a specific calculated value (e.g., 30) exists in the markup. | Explain how HtmlSaveOptions can be configured to export only evaluated results while preserving cell styling.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlFormulaDemo
{
    // Demonstrates how to create a workbook, add numeric data and a SUM formula, configure HtmlSaveOptions to calculate formulas and omit the original expression, save as HTML, and programmatically confirm that the computed value (30) appears in the generated markup.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            // Add a formula that sums A1 and A2
            sheet.Cells["A3"].Formula = "=SUM(A1:A2)";

            // Configure HTML save options to calculate formulas before saving
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                CalculateFormula = true,
                ExportFormula = false // export only the calculated values
            };

            // Save the workbook as HTML
            string htmlPath = "WorkbookWithCalculatedFormulas.html";
            workbook.Save(htmlPath, htmlOptions);

            // Verify that the calculated result (30) appears in the generated HTML
            string htmlContent = File.ReadAllText(htmlPath);
            bool containsResult = htmlContent.Contains(">30<");
            Console.WriteLine($"HTML contains calculated result: {containsResult}");
        }
    }
}
