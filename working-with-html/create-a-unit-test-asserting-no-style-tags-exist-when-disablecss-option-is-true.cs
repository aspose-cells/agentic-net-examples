// Title: C# Unit Test: Verify HtmlSaveOptions.DisableCss Prevents <style> Tags in Aspose.Cells HTML Export
// Description: Creates a workbook, applies bold red formatting to cell A1, saves it as HTML with HtmlSaveOptions.DisableCss enabled, and asserts that the resulting markup contains no <style> elements.
// Keywords: Aspose.Cells | HtmlSaveOptions.DisableCss | C# unit test | HTML export without CSS | inline styles only | assert no style tag | Aspose.Cells .NET testing
// Common Searches: Aspose.Cells unit test for DisableCss | how to assert no <style> tag in exported HTML | C# test HtmlSaveOptions.DisableCss behavior | verify inline styling only Aspose.Cells
// Developer Intent: Confirm that setting HtmlSaveOptions.DisableCss to true eliminates all <style> blocks from the generated HTML.
// Use Cases: Automated CI validation that Excel‑to‑HTML conversion uses only inline styles. | Generating email‑ready HTML from spreadsheets without external CSS. | Regression testing to detect unintended CSS output after library updates.
// AI Prompts: Create an MSTest method that saves a workbook to HTML with DisableCss=true and asserts the output lacks <style> tags. | Write an xUnit test for Aspose.Cells that verifies HtmlSaveOptions.DisableCss removes all style elements. | Provide a NUnit example that checks for the absence of <style> blocks when exporting a workbook to HTML using Aspose.Cells.

using System;
using System.IO;
using System.Text;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Creates a workbook, applies bold red formatting to cell A1, saves it as HTML with HtmlSaveOptions.DisableCss enabled, and asserts that the resulting markup contains no <style> elements.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and apply formatting to cell A1
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Cells["A1"].PutValue("Styled Text");
                Style style = worksheet.Cells["A1"].GetStyle();
                style.Font.IsBold = true;
                style.Font.Color = Color.Red;
                worksheet.Cells["A1"].SetStyle(style);

                // Configure HtmlSaveOptions to disable CSS (use only inline styles)
                HtmlSaveOptions options = new HtmlSaveOptions
                {
                    DisableCss = true
                };

                // Save the workbook to a memory stream using the options
                using (MemoryStream stream = new MemoryStream())
                {
                    workbook.Save(stream, options);
                    stream.Position = 0;
                    string htmlContent = new StreamReader(stream, Encoding.UTF8).ReadToEnd();

                    // Verify that the generated HTML does not contain any <style> tags
                    bool containsStyleTag = htmlContent.IndexOf("<style", StringComparison.OrdinalIgnoreCase) >= 0;
                    if (containsStyleTag)
                    {
                        Console.WriteLine("Test Failed: HTML output contains a <style> tag despite DisableCss being true.");
                    }
                    else
                    {
                        Console.WriteLine("Test Passed: No <style> tag found in HTML output.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception occurred: {ex.Message}");
            }
        }
    }
}
