// Title: Aspose.Cells for .NET: Export Workbook to HTML with Calculated Formulas
// Description: Creates a workbook, fills cells A1 and A2, adds a SUM formula in A3, sets HtmlSaveOptions.CalculateFormula to true, saves the file as HTML, and verifies that the computed value (30) appears in the generated markup.
// Keywords: Aspose.Cells | .NET | C# | HTML export | formula evaluation | HtmlSaveOptions | CalculateFormula | workbook to HTML | SUM formula | automated verification
// Common Searches: Aspose.Cells export HTML with formula results | C# save Excel as HTML with calculated values | HtmlSaveOptions CalculateFormula example | verify formula result in exported HTML Aspose | how to include evaluated formulas in HTML using Aspose.Cells
// Developer Intent: Generate an HTML representation of an Excel workbook where all formulas are evaluated beforehand, and optionally confirm that the expected numeric outcomes are present in the output.
// Use Cases: Publish web‑ready reports that display totals and percentages already calculated. | Create documentation from spreadsheets where formulas must be resolved before publishing. | Automate QA tests that ensure exported HTML contains specific calculated values. | Integrate Excel‑to‑HTML conversion into a .NET web application with pre‑computed data.
// AI Prompts: Show how to add custom CSS to cells that contain evaluated results during HTML export with Aspose.Cells. | Provide code to export multiple worksheets into a single HTML file while preserving each sheet's calculated values. | Explain how to log every formula's computed result during the HTML conversion process for debugging.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills cells A1 and A2, adds a SUM formula in A3, sets HtmlSaveOptions.CalculateFormula to true, saves the file as HTML, and verifies that the computed value (30) appears in the generated markup.
    public class HtmlWithCalculatedFormulasDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate cells with values
                worksheet.Cells["A1"].PutValue(10);
                worksheet.Cells["A2"].PutValue(20);

                // Add a formula that sums A1 and A2
                worksheet.Cells["A3"].Formula = "=SUM(A1:A2)";

                // Configure HtmlSaveOptions to calculate formulas before saving
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    CalculateFormula = true // Ensure formulas are evaluated
                };

                // Save the workbook as HTML; the result of A3 will be written instead of the formula text
                string htmlPath = "HtmlWithCalculatedFormulas.html";
                workbook.Save(htmlPath, saveOptions);

                // Optional verification: read the generated HTML and check that the calculated value (30) appears
                if (File.Exists(htmlPath))
                {
                    string htmlContent = File.ReadAllText(htmlPath);
                    bool containsResult = htmlContent.Contains("30");
                    Console.WriteLine($"HTML contains calculated result (30): {containsResult}");
                }
                else
                {
                    Console.WriteLine($"Failed to generate HTML file at path: {htmlPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in HtmlWithCalculatedFormulasDemo: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                HtmlWithCalculatedFormulasDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
