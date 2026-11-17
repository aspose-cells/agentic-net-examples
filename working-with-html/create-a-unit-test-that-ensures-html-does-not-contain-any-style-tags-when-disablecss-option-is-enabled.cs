// Title: C# unit test using Aspose.Cells to confirm HTML output has no <style> tags when HtmlSaveOptions.DisableCss is true
// Common Searches: how to write a C# unit test for Aspose.Cells HTML export with DisableCss option | Aspose.Cells HtmlSaveOptions.DisableCss unit test example in .NET | verify that generated HTML from Aspose.Cells does not contain style elements | C# test to ensure Aspose.Cells HTML output excludes CSS when disabled | unit testing Aspose.Cells HTML save without embedded style tags
// Tags: Aspose.Cells HtmlSaveOptions.DisableCss verification | C# Aspose.Cells HTML export unit test | HTML output without <style> tags Aspose.Cells | disable CSS generation in Aspose.Cells HTML save | assert no style elements in generated HTML C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTests
{
    // The example creates a workbook, populates cells, saves it to HTML with HtmlSaveOptions.DisableCss set to true, reads the HTML from a memory stream, and asserts that the output does not contain any <style> tags, indicating that CSS generation is correctly disabled.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and add sample data
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Header");
                sheet.Cells["A2"].PutValue(123);
                sheet.Cells["B1"].PutValue("Date");
                sheet.Cells["B2"].PutValue(DateTime.Now);

                // Configure HTML save options to disable CSS generation
                var htmlOptions = new HtmlSaveOptions
                {
                    DisableCss = true,
                    ExportActiveWorksheetOnly = true
                };

                // Save the workbook to a memory stream as HTML
                using (var ms = new MemoryStream())
                {
                    workbook.Save(ms, htmlOptions);
                    ms.Position = 0;

                    // Read the generated HTML as a string
                    string htmlContent;
                    using (var reader = new StreamReader(ms))
                    {
                        htmlContent = reader.ReadToEnd();
                    }

                    // Verify that no <style> tags are present in the output
                    if (htmlContent.Contains("<style", StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine("Test Failed: HTML output contains <style> tags when DisableCss is enabled.");
                    }
                    else
                    {
                        Console.WriteLine("Test Passed: HTML output does not contain <style> tags.");
                    }
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
