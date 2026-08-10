// Title: Trim Whitespace After <br> Tags with HtmlLoadOptions.DeleteRedundantSpaces in Aspose.Cells (C#)
// Description: Demonstrates how to enable HtmlLoadOptions.DeleteRedundantSpaces to remove spaces that follow <br> line‑break tags when loading HTML into an Aspose.Cells workbook. The example builds an HTML string, configures the option, loads the content via a MemoryStream, shows the cleaned cell value, and saves the workbook.
// Keywords: Aspose.Cells | HtmlLoadOptions | DeleteRedundantSpaces | C# | trim whitespace | HTML to Excel | remove extra spaces | line break handling | memory stream import | Excel export
// Common Searches: Aspose.Cells remove spaces after <br> tag | HtmlLoadOptions DeleteRedundantSpaces C# example | trim whitespace when importing HTML to Excel | load HTML string into workbook without extra spaces | Aspose.Cells HTML import whitespace trimming
// Developer Intent: Enable the DeleteRedundantSpaces flag so that any spaces after <br> tags are automatically stripped during HTML‑to‑Excel conversion.
// Use Cases: Import HTML reports containing line‑break tags while preserving clean cell text. | Automate conversion of web‑generated tables to Excel without unwanted padding. | Prepare Excel files for downstream processing where extra spaces cause parsing errors.
// AI Prompts: Show how to use HtmlLoadOptions.DeleteRedundantSpaces in C# to trim spaces after <br> when loading HTML into an Aspose.Cells workbook. | Provide a C# code snippet that reads an HTML string, enables whitespace trimming, and saves the result as an Excel file with Aspose.Cells. | Explain the impact of the DeleteRedundantSpaces option on cell values after importing HTML with Aspose.Cells.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to enable HtmlLoadOptions.DeleteRedundantSpaces to remove spaces that follow <br> line‑break tags when loading HTML into an Aspose.Cells workbook. The example builds an HTML string, configures the option, loads the content via a MemoryStream, shows the cleaned cell value, and saves the workbook.
    public class HtmlLoadOptionsDeleteRedundantSpacesDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public static void Run()
        {
            // HTML with line breaks (<br>) followed by extra spaces
            string html = "<p>Line1<br>   Line2<br>    Line3</p>";

            // Configure HtmlLoadOptions to delete redundant spaces after line breaks
            HtmlLoadOptions loadOptions = new HtmlLoadOptions
            {
                DeleteRedundantSpaces = true // Enable whitespace trimming
            };

            // Convert HTML string to a memory stream
            byte[] htmlBytes = Encoding.UTF8.GetBytes(html);
            using (MemoryStream stream = new MemoryStream(htmlBytes))
            {
                // Load the workbook using the configured options
                Workbook workbook = new Workbook(stream, loadOptions);

                // Access the first worksheet and first cell (A1) where the text is placed
                Worksheet worksheet = workbook.Worksheets[0];
                Cell cell = worksheet.Cells["A1"];

                // Output the cell text after redundant spaces have been removed
                Console.WriteLine("Cell text after trimming spaces: " + cell.StringValue);

                // Save the workbook to verify the result (optional)
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TrimmedHtmlOutput.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + outputPath);
            }
        }
    }
}
