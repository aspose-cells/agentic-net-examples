// Title: C# Unit Test for Verifying Cell Value and <td id> in Aspose.Cells HTML Export
// Description: Shows how to create a workbook, set A1 to "Aspose Test", export to HTML with HtmlSaveOptions.CellNameAttribute = "id", read the output, and assert that the HTML contains the text and a <td id='A1'> element. Includes temporary file handling and cleanup.
// Keywords: Aspose.Cells | C# | .NET | HTML export | HtmlSaveOptions | CellNameAttribute | unit test | MSTest | xUnit | NUnit | regex validation | temporary file | CI | regression testing
// Common Searches: Aspose.Cells unit test HTML export | verify cell value in exported HTML C# | check td id attribute Aspose.Cells | how to assert HTML output with Aspose.Cells | C# test for HtmlSaveOptions CellNameAttribute | automated test for Aspose.Cells HTML export
// Developer Intent: Confirm that the HTML produced by Aspose.Cells includes the expected cell content and the correct id attribute for that cell.
// Use Cases: Continuous integration validation of HTML export | Regression testing after Aspose.Cells version upgrades | Automated verification of custom cell identifier mapping | Quality gate for documentation generation pipelines | Reference sample for developers writing Aspose.Cells unit tests
// AI Prompts: Generate an MSTest method that creates a workbook, writes "Aspose Test" to A1, saves as HTML with CellNameAttribute='id', reads the file, and asserts the presence of the text and <td id='A1'>, including setup and teardown. | Provide an xUnit test example for Aspose.Cells that checks the exported HTML for a specific cell value and its id attribute using regular expressions and ensures temporary files are deleted. | Write a NUnit test that validates Aspose.Cells HTML export by confirming the cell value and <td id> attribute, with proper file cleanup and exception handling.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Shows how to create a workbook, set A1 to "Aspose Test", export to HTML with HtmlSaveOptions.CellNameAttribute = "id", read the output, and assert that the HTML contains the text and a <td id='A1'> element. Includes temporary file handling and cleanup.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                HtmlExportDemo.Run();
                Console.WriteLine("HTML export test completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public static class HtmlExportDemo
    {
        public static void Run()
        {
            // Create a new workbook and set a value in cell A1
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Aspose Test");

            // Configure HTML save options to include cell identifiers
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                CellNameAttribute = "id" // <td id='A1'>...</td>
            };

            // Use a temporary file for the HTML output
            string tempHtmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".html");

            try
            {
                // Save the workbook as HTML
                workbook.Save(tempHtmlPath, saveOptions);

                // Ensure the file was created before reading
                if (!File.Exists(tempHtmlPath))
                {
                    throw new FileNotFoundException("HTML output file was not created.", tempHtmlPath);
                }

                // Read the generated HTML content
                string htmlContent = File.ReadAllText(tempHtmlPath);

                // Verify that the HTML contains the expected cell value
                bool containsValue = Regex.IsMatch(htmlContent, @"Aspose Test", RegexOptions.IgnoreCase);
                if (!containsValue)
                {
                    throw new InvalidOperationException("The HTML output does not contain the expected cell value.");
                }

                // Verify that the cell identifier attribute is present
                bool containsCellId = Regex.IsMatch(htmlContent, @"<td\s+id=['""]A1['""]", RegexOptions.IgnoreCase);
                if (!containsCellId)
                {
                    throw new InvalidOperationException("The HTML output does not contain the expected cell ID attribute.");
                }
            }
            finally
            {
                // Clean up the temporary file
                if (File.Exists(tempHtmlPath))
                {
                    try
                    {
                        File.Delete(tempHtmlPath);
                    }
                    catch
                    {
                        // Suppress any exceptions during cleanup
                    }
                }
            }
        }
    }
}
