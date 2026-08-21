// Title: C# – Compare HTML loading with Aspose.Cells DeleteRedundantSpaces option
// Description: Loads the same HTML string into two Aspose.Cells workbooks – first with the default HtmlLoadOptions (DeleteRedundantSpaces = false) and then with DeleteRedundantSpaces enabled – extracts cell A1 text, prints both values, checks equality, and optionally saves each workbook for visual inspection.
// Keywords: Aspose.Cells DeleteRedundantSpaces | HtmlLoadOptions C# | load HTML Aspose.Cells | trim redundant spaces Aspose | compare HTML whitespace Aspose.Cells
// Common Searches: Aspose.Cells enable DeleteRedundantSpaces | HTML whitespace handling Aspose.Cells | difference DeleteRedundantSpaces true false | compare cell values after trimming spaces | how to remove extra spaces when loading HTML
// Developer Intent: Understand how the DeleteRedundantSpaces flag affects text extracted from HTML when imported into an Aspose.Cells workbook.
// Use Cases: Validate whether whitespace trimming changes imported text for downstream processing. | Generate side‑by‑side Excel files to visually assess layout impact of space removal. | Automate a decision‑making step that selects the appropriate HtmlLoadOptions setting based on content comparison.
// AI Prompts: Provide a C# example that loads an HTML file twice with Aspose.Cells – once with DeleteRedundantSpaces disabled and once enabled – and outputs the cell A1 strings. | Explain how DeleteRedundantSpaces influences cell string values during HTML import and list scenarios where enabling it improves data quality. | Write a C# unit test that asserts the cell text differs when HtmlLoadOptions.DeleteRedundantSpaces is true versus false.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsHtmlSpaceComparison
{
    // Loads the same HTML string into two Aspose.Cells workbooks – first with the default HtmlLoadOptions (DeleteRedundantSpaces = false) and then with DeleteRedundantSpaces enabled – extracts cell A1 text, prints both values, checks equality, and optionally saves each workbook for visual inspection.
    class Program
    {
        static void Main()
        {
            // Sample HTML containing redundant spaces.
            string html = "<p>   This    text   has   redundant   spaces   </p>";

            // Convert the HTML string to a UTF‑8 byte array and wrap it in a MemoryStream.
            byte[] htmlBytes = Encoding.UTF8.GetBytes(html);
            using (MemoryStream stream = new MemoryStream(htmlBytes))
            {
                // -------------------------------------------------
                // Load without deleting redundant spaces (default).
                // -------------------------------------------------
                Workbook wbDefault = new Workbook(stream); // uses default HtmlLoadOptions (DeleteRedundantSpaces = false)
                string textDefault = wbDefault.Worksheets[0].Cells["A1"].StringValue;
                Console.WriteLine("Text without DeleteRedundantSpaces: \"" + textDefault + "\"");

                // Reset the stream position for the second load.
                stream.Position = 0;

                // -------------------------------------------------
                // Load with DeleteRedundantSpaces enabled.
                // -------------------------------------------------
                HtmlLoadOptions loadOpts = new HtmlLoadOptions
                {
                    DeleteRedundantSpaces = true
                };
                Workbook wbTrimmed = new Workbook(stream, loadOpts);
                string textTrimmed = wbTrimmed.Worksheets[0].Cells["A1"].StringValue;
                Console.WriteLine("Text with DeleteRedundantSpaces:    \"" + textTrimmed + "\"");

                // -------------------------------------------------
                // Simple comparison output.
                // -------------------------------------------------
                bool areEqual = string.Equals(textDefault, textTrimmed, StringComparison.Ordinal);
                Console.WriteLine("Are the texts equal? " + areEqual);

                // -------------------------------------------------
                // Save both workbooks to Excel files for visual inspection (optional).
                // -------------------------------------------------
                wbDefault.Save("WithoutTrim.xlsx");
                wbTrimmed.Save("WithTrim.xlsx");
                Console.WriteLine("Workbooks saved as 'WithoutTrim.xlsx' and 'WithTrim.xlsx'.");
            }
        }
    }
}
