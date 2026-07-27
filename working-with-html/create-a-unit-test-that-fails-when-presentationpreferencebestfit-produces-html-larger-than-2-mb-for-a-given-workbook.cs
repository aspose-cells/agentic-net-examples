// Title: Aspose.Cells .NET unit test: fail when PresentationPreference generates HTML larger than 2 MB
// Description: The example builds a workbook with 5,000 rows and 50 columns of 200‑character strings, saves it to a MemoryStream using HtmlSaveOptions (PresentationPreference enabled, SaveAsSingleFile true), and asserts that the resulting HTML size does not exceed 2 MB, reporting success or failure.
// Keywords: Aspose.Cells | HtmlSaveOptions | PresentationPreference | BestFit | HTML export size | unit test | C# | .NET | SaveAsSingleFile | memory stream | CI regression | file size limit
// Common Searches: Aspose.Cells unit test HTML size limit | PresentationPreference exceeds 2 MB | assert generated HTML size in .NET | HtmlSaveOptions SaveAsSingleFile size check | how to fail test when HTML > 2 MB Aspose.Cells
// Developer Intent: Create an automated test that fails if enabling PresentationPreference produces an HTML file larger than 2 MB.
// Use Cases: Validate that HTML export stays within bandwidth constraints for web delivery. | Detect size‑regression bugs in CI pipelines when workbook data or save options change. | Ensure single‑file HTML output consolidates resources without breaching a predefined size threshold.
// AI Prompts: Generate an MSTest method that creates a 5,000 × 50 workbook, saves it to HTML with PresentationPreference = true and SaveAsSingleFile = true, and asserts the MemoryStream length is under 2 MB. | Write a NUnit test that verifies Aspose.Cells HtmlSaveOptions with PresentationPreference enabled does not exceed a 2 MB limit for a large workbook, failing with a clear message if it does. | Provide a xUnit test example that builds a workbook, fills it with repetitive data, exports to HTML using HtmlSaveOptions, and uses Assert.True to enforce a maximum file size of 2 MB.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // The example builds a workbook with 5,000 rows and 50 columns of 200‑character strings, saves it to a MemoryStream using HtmlSaveOptions (PresentationPreference enabled, SaveAsSingleFile true), and asserts that the resulting HTML size does not exceed 2 MB, reporting success or failure.
    public class HtmlPresentationPreferenceTest
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and populate it with a large amount of data.
                using (Workbook workbook = new Workbook())
                {
                    Worksheet sheet = workbook.Worksheets[0];

                    // Fill 5000 rows and 50 columns with repetitive text.
                    for (int row = 0; row < 5000; row++)
                    {
                        for (int col = 0; col < 50; col++)
                        {
                            sheet.Cells[row, col].PutValue(new string('A', 200));
                        }
                    }

                    // Configure HTML save options with PresentationPreference enabled.
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                    {
                        PresentationPreference = true, // Enable more beautiful presentation.
                        SaveAsSingleFile = true          // Consolidate output into a single HTML file.
                    };

                    // Save the workbook to a memory stream.
                    using (MemoryStream htmlStream = new MemoryStream())
                    {
                        workbook.Save(htmlStream, htmlOptions);

                        long htmlSizeInBytes = htmlStream.Length;
                        const long TwoMegabytes = 2L * 1024 * 1024;

                        // Evaluate the generated HTML size.
                        if (htmlSizeInBytes <= TwoMegabytes)
                        {
                            Console.WriteLine($"Success: HTML size {htmlSizeInBytes} bytes is within the 2 MB limit.");
                        }
                        else
                        {
                            Console.WriteLine($"Failure: HTML size {htmlSizeInBytes} bytes exceeds the 2 MB limit.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Runtime safety: report any unexpected errors.
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
