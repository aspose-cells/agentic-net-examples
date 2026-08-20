// Title: C# Unit Test to Verify a Cell Value in Aspose.Cells HTML Export
// Description: Creates a workbook, writes "HelloWorld" to cell A1, saves it as HTML with HtmlSaveOptions (CellNameAttribute = "id"), reads the output file, and uses a regular expression to assert that the expected value appears inside a <td> element.
// Keywords: Aspose.Cells HTML export unit test | C# verify cell value in HTML | Aspose.Cells HtmlSaveOptions id attribute | regex validation of generated HTML | automated test for workbook to HTML conversion
// Common Searches: unit test Aspose.Cells HTML output contains specific text | C# assert cell A1 value in exported HTML | how to validate Aspose.Cells HTML export with regex | test workbook to HTML conversion Aspose.Cells | verify cell content after saving as HTML in C#
// Developer Intent: Write an automated test that confirms a known worksheet cell value is present in the HTML file produced by Aspose.Cells.
// Use Cases: Regression testing to ensure data integrity after HTML conversion | CI/CD pipeline check that critical cell values survive export | Validation of custom HtmlSaveOptions such as CellNameAttribute
// AI Prompts: Generate an MSTest method that sets "HelloWorld" in A1, saves the workbook as HTML with Aspose.Cells, and asserts the HTML contains a <td> with that text. | Create an xUnit test that exports a workbook to HTML using HtmlSaveOptions and verifies the expected cell value using a regular expression. | Provide a NUnit example that writes a value to a worksheet, converts it to HTML, and checks for the value inside a table cell.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Creates a workbook, writes "HelloWorld" to cell A1, saves it as HTML with HtmlSaveOptions (CellNameAttribute = "id"), reads the output file, and uses a regular expression to assert that the expected value appears inside a <td> element.
    class Program
    {
        private const string OutputHtmlPath = "TestOutput.html";

        static void Main()
        {
            try
            {
                // Ensure previous output does not interfere with the run
                if (File.Exists(OutputHtmlPath))
                {
                    File.Delete(OutputHtmlPath);
                }

                // Create a new workbook and access the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Put a known value into cell A1
                const string expectedValue = "HelloWorld";
                worksheet.Cells["A1"].PutValue(expectedValue);

                // Save the workbook as HTML with cell name attribute for easier verification
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    CellNameAttribute = "id"
                };
                workbook.Save(OutputHtmlPath, saveOptions);

                // Verify that the HTML file was created
                if (!File.Exists(OutputHtmlPath))
                {
                    Console.WriteLine($"Failed to create HTML file at '{OutputHtmlPath}'.");
                    return;
                }

                // Read the generated HTML content
                string htmlContent = File.ReadAllText(OutputHtmlPath);

                // Check that the HTML contains the expected cell value inside a <td> element
                bool containsValue = Regex.IsMatch(
                    htmlContent,
                    $@"<td[^>]*>\s*{Regex.Escape(expectedValue)}\s*</td>",
                    RegexOptions.IgnoreCase);

                if (containsValue)
                {
                    Console.WriteLine("Success: HTML output contains the expected cell value.");
                }
                else
                {
                    Console.WriteLine($"Failure: HTML output does not contain the expected cell value '{expectedValue}'.");
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
