// Title: C# Aspose.Cells: Export workbook with formulas to HTML and confirm calculated total
// Common Searches: how to export a workbook with formulas to HTML using Aspose.Cells in C# | Aspose.Cells calculate formulas before saving as HTML C# example | verify that a SUM formula result appears in generated HTML file Aspose.Cells | C# code to read generated HTML and check calculated cell value from Aspose.Cells | Aspose.Cells SaveFormat.Html preserving evaluated formulas C#
// Tags: Aspose.Cells export to HTML with pre‑calculated formulas | C# workbook formula evaluation prior to HTML save | assert calculated cell value in Aspose.Cells HTML output | SaveFormat.Html usage with SUM formula in C# | read and validate generated HTML content Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlFormulaDemo
{
    // The sample creates a Workbook, fills it with item data, adds a SUM formula, forces formula calculation with CalculateFormula(), saves the workbook as HTML, reads the generated HTML file, and checks that the evaluated total (4.3) is present, outputting a pass/fail verification result.
    class Program
    {
        static void Main(string[] args)
        {
            // Define file paths (relative to the executable directory)
            string htmlOutputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "FormulaDemo.html");

            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // 2. Populate cells with sample data
                Cells cells = worksheet.Cells;
                cells["A1"].PutValue("Item");
                cells["B1"].PutValue("Price");
                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue(2.5);
                cells["A3"].PutValue("Orange");
                cells["B3"].PutValue(1.8);
                cells["A4"].PutValue("Total:");
                // 3. Set a formula to calculate the sum of the prices
                cells["B4"].Formula = "=SUM(B2:B3)";

                // 4. Force calculation of all formulas in the workbook
                workbook.CalculateFormula();

                // Capture the calculated result for verification
                double calculatedTotal = cells["B4"].DoubleValue; // Expected: 4.3

                // 5. Save the workbook as HTML
                workbook.Save(htmlOutputPath, SaveFormat.Html);

                // 6. Read the generated HTML file
                string htmlContent = File.ReadAllText(htmlOutputPath);

                // 7. Verify that the calculated result appears in the HTML
                // Use invariant culture formatting to match the HTML representation
                string expectedValueString = calculatedTotal.ToString(System.Globalization.CultureInfo.InvariantCulture);
                bool containsResult = htmlContent.Contains(expectedValueString);

                // 8. Output verification result
                Console.WriteLine("HTML file generated at: " + htmlOutputPath);
                Console.WriteLine("Calculated total (B4): " + calculatedTotal);
                Console.WriteLine("Verification: " + (containsResult ? "PASS" : "FAIL"));
                if (!containsResult)
                {
                    Console.WriteLine("Expected value \"" + expectedValueString + "\" was not found in the HTML output.");
                }
            }
            catch (Exception ex)
            {
                // Report any unexpected errors
                Console.WriteLine("An error occurred: " + ex.Message);
                Console.WriteLine("Stack Trace:");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
