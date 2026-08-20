// Title: Validate plain‑text rendering of large numbers in HTML export with Aspose.Cells for .NET
// Description: Shows how to place a 16‑digit value as a string, configure SignificantDigitsType to Digits15, save the workbook as HTML, and programmatically verify that the HTML contains the exact number without scientific (E+) notation.
// Keywords: Aspose.Cells | .NET | C# | HTML export | large integer | plain text output | scientific notation | SignificantDigitsType | Digits15 | numeric formatting | automated verification
// Common Searches: Aspose.Cells keep large numbers from becoming scientific notation in HTML | C# export workbook to HTML without exponential format | SignificantDigitsType Digits15 example | verify exact cell value in generated HTML Aspose.Cells | prevent scientific notation when saving as HTML .NET
// Developer Intent: Ensure the HTML produced by Aspose.Cells displays a large integer exactly as entered, not in exponential form.
// Use Cases: Export financial statements containing 16‑digit account numbers to HTML while preserving every digit. | Create an automated test that reads the saved HTML file and asserts the presence of the original number string and the absence of "E+". | Apply a workbook‑wide setting to force plain numeric formatting for all cells before HTML conversion.
// AI Prompts: Generate C# code using Aspose.Cells that saves a worksheet as HTML and checks that a 16‑digit number appears unchanged in the output. | Write an MSTest unit test that sets SignificantDigitsType to Digits15, exports to HTML, and asserts no scientific notation is present. | Explain the effect of SignificantDigitsType.Digits15 on number formatting during HTML export in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLargeNumberHtmlDemo
{
    // Shows how to place a 16‑digit value as a string, configure SignificantDigitsType to Digits15, save the workbook as HTML, and programmatically verify that the HTML contains the exact number without scientific (E+) notation.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Large number that could be displayed in exponential form by default
            const string largeNumber = "1234567890123456";

            // Put the large number as a string to preserve exact representation
            // (If PutValue with a numeric type is used, Excel may format it automatically)
            sheet.Cells["A1"].PutValue(largeNumber);

            // Ensure the workbook uses plain numeric formatting (no scientific notation)
            // Set the significant digits type to 15 digits which forces plain text output
            workbook.Settings.SignificantDigitsType = SignificantDigitsType.Digits15;

            // Save the workbook as HTML
            string htmlPath = "LargeNumberOutput.html";
            workbook.Save(htmlPath, SaveFormat.Html);

            // Read the generated HTML file
            string htmlContent = File.ReadAllText(htmlPath);

            // Verify that the large number appears as plain text, not in exponential form
            bool containsPlain = htmlContent.Contains(largeNumber);
            bool containsExponential = htmlContent.Contains("E+");

            Console.WriteLine("HTML contains plain large number: " + containsPlain);
            Console.WriteLine("HTML contains exponential representation: " + containsExponential);

            // Simple assertion output
            if (containsPlain && !containsExponential)
            {
                Console.WriteLine("Verification succeeded: Large numbers are displayed as plain text.");
            }
            else
            {
                Console.WriteLine("Verification failed: Large numbers are displayed in exponential form.");
            }
        }
    }
}
