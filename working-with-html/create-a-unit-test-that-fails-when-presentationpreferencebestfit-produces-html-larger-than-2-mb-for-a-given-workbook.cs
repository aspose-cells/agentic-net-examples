// Title: C# unit test to ensure Aspose.Cells HTML output with PresentationPreference exceeds 2 MB
// Description: Creates a workbook with 5,000 rows × 50 columns of 200‑character strings, saves it as a single HTML file using HtmlSaveOptions (PresentationPreference = true), measures the resulting stream size, and fails the test when the HTML is 2 MB or smaller.
// Keywords: Aspose.Cells | C# | HtmlSaveOptions | PresentationPreference | HTML size test | 2 MB limit | unit test | SaveAsSingleFile | MemoryStream | performance regression
// Common Searches: Aspose.Cells unit test HTML size | PresentationPreference HTML output size | C# check HTML file size Aspose.Cells | Validate HtmlSaveOptions SaveAsSingleFile size | Measure Aspose.Cells HTML stream length
// Developer Intent: Write an automated test that fails if the HTML generated with PresentationPreference does not exceed 2 MB.
// Use Cases: Detect regressions where PresentationPreference reduces HTML payload | Integrate size validation into CI pipelines for large workbook exports | Confirm that SaveAsSingleFile aggregates resources into a single HTML document of expected size
// AI Prompts: Generate an MSTest method that asserts HtmlSaveOptions.PresentationPreference produces HTML larger than 2 MB for a workbook filled with 5,000 rows and 50 columns of 200‑character strings. | Create a NUnit test that saves a workbook to a MemoryStream with SaveAsSingleFile = true, checks stream.Length, and fails when size ≤ 2 MB. | Provide an xUnit test example that captures HTML output, measures its length, and throws an AssertionFailedException if the size does not exceed 2 MB.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsTests
{
    // Creates a workbook with 5,000 rows × 50 columns of 200‑character strings, saves it as a single HTML file using HtmlSaveOptions (PresentationPreference = true), measures the resulting stream size, and fails the test when the HTML is 2 MB or smaller.
    class Program
    {
        static void Main()
        {
            try
            {
                RunHtmlSizeTest();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        static void RunHtmlSizeTest()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with a large amount of data to increase HTML size
            const int rows = 5000;
            const int cols = 50;
            string longText = new string('x', 200); // 200 characters per cell

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    sheet.Cells[r, c].PutValue(longText);
                }
            }

            // Configure HTML save options with PresentationPreference enabled
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                PresentationPreference = true, // more beautiful presentation
                SaveAsSingleFile = true          // generate a single HTML file for size measurement
            };

            // Save the workbook to a memory stream using the configured options
            using (MemoryStream htmlStream = new MemoryStream())
            {
                workbook.Save(htmlStream, htmlOptions);

                long htmlSize = htmlStream.Length;
                const long twoMegabytes = 2L * 1024 * 1024;

                if (htmlSize > twoMegabytes)
                {
                    Console.WriteLine($"Success: HTML size {htmlSize} bytes exceeds 2 MB.");
                }
                else
                {
                    Console.WriteLine($"Failure: HTML size {htmlSize} bytes does not exceed 2 MB.");
                }
            }
        }
    }
}
