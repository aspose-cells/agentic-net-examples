// Title: C# Unit Test: Verify HtmlSaveOptions.DisableCss Generates HTML Without <style> Tags in Aspose.Cells
// Description: Creates a workbook, applies bold red formatting, saves it to HTML with HtmlSaveOptions.DisableCss enabled, and asserts that the output contains no <style> elements while the formatting appears as inline style attributes.
// Keywords: Aspose.Cells | HtmlSaveOptions | DisableCss | C# unit test | .NET HTML export | inline styles | no style tag | automated testing | CI validation | Aspose.Cells HTML output
// Common Searches: Aspose.Cells unit test for DisableCss | how to ensure HtmlSaveOptions.DisableCss removes style tags | C# test HTML export without CSS using Aspose.Cells | verify inline styles in Aspose.Cells HTML output | disable CSS in Aspose.Cells HTML export unit test
// Developer Intent: Confirm that setting HtmlSaveOptions.DisableCss to true prevents <style> blocks from being emitted and forces formatting to be rendered as inline CSS.
// Use Cases: Add a CI check that HTML exported from Aspose.Cells complies with email clients that block embedded CSS. | Guard against regressions where future library updates might re‑introduce <style> sections when DisableCss is used. | Ensure that workbook formatting (e.g., bold, color) is correctly translated into inline style attributes for downstream processing.
// AI Prompts: Generate an MSTest method that creates a workbook, applies bold red formatting, saves to HTML with HtmlSaveOptions.DisableCss, and asserts the absence of <style> tags and the presence of inline style="font-weight:bold". | Write a xUnit test that loads a workbook, configures HtmlSaveOptions.DisableCss, writes to a MemoryStream, reads the HTML string, and verifies no style blocks exist while inline CSS reflects the cell style. | Provide a NUnit example that checks Aspose.Cells HTML export for inline styling only when the DisableCss option is enabled, including assertions for missing <style> elements and correct inline attributes.

using System;
using System.IO;
using System.Text;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Creates a workbook, applies bold red formatting, saves it to HTML with HtmlSaveOptions.DisableCss enabled, and asserts that the output contains no <style> elements while the formatting appears as inline style attributes.
    public class HtmlSaveOptionsTests
    {
        public void Run()
        {
            try
            {
                // Create a new workbook and add some formatted data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cell cell = sheet.Cells["A1"];
                cell.PutValue("Styled Text");

                // Apply a style that would normally be emitted as CSS
                Style style = cell.GetStyle();
                style.Font.IsBold = true;
                style.Font.Color = Color.Red;
                cell.SetStyle(style);

                // Configure HtmlSaveOptions to disable CSS (use only inline styles)
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    DisableCss = true
                };

                // Save the workbook to a memory stream as HTML
                using (MemoryStream htmlStream = new MemoryStream())
                {
                    workbook.Save(htmlStream, saveOptions);
                    htmlStream.Position = 0;

                    // Read the generated HTML as a string
                    string htmlContent;
                    using (StreamReader reader = new StreamReader(htmlStream, Encoding.UTF8))
                    {
                        htmlContent = reader.ReadToEnd();
                    }

                    // Verify that no <style> tags are present in the output
                    if (htmlContent.IndexOf("<style", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        Console.WriteLine("FAIL: HTML output contains <style> tags even though DisableCss is enabled.");
                    }
                    else
                    {
                        Console.WriteLine("PASS: No <style> tags found.");
                    }

                    // Ensure that inline style attributes are present (formatting applied)
                    if (htmlContent.Contains("style=\"") && htmlContent.Contains("font-weight:bold"))
                    {
                        Console.WriteLine("PASS: Inline style attributes are present.");
                    }
                    else
                    {
                        Console.WriteLine("FAIL: Inline style attributes are missing; formatting may not have been applied.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main()
        {
            HtmlSaveOptionsTests test = new HtmlSaveOptionsTests();
            test.Run();
        }
    }
}
