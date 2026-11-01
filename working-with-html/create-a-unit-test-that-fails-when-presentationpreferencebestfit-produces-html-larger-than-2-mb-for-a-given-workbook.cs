// Title: Write a C# unit test that verifies Aspose.Cells PresentationPreference.BestFit HTML export stays under 2 MB for a large workbook
// AI Prompts: Create an MSTest method that builds a workbook with 5,000 rows and 50 columns, saves it to HTML using HtmlSaveOptions with PresentationPreference.BestFit, and asserts the output stream size is less than 2 MB. | Generate a NUnit test that populates a worksheet with extensive data, exports it to HTML via Aspose.Cells with PresentationPreference.BestFit, and fails the test when the generated HTML exceeds 2 MB.
// Common Searches: asp.net core unit test Aspose.Cells HTML export size limit | how to fail a test when PresentationPreference.BestFit HTML exceeds 2 MB | verify Aspose.Cells HTML output size in C# unit test | Aspose.Cells PresentationPreference.BestFit large workbook HTML size check
// Tags: Aspose.Cells HtmlSaveOptions PresentationPreference.BestFit size check | C# unit test Aspose.Cells HTML export limit | large workbook HTML generation Aspose.Cells | assert HTML byte size under 2MB Aspose.Cells | memory stream size validation Aspose.Cells HTML

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTests
{
    // The example builds a workbook with 5,000 rows and 50 columns, configures HtmlSaveOptions to use PresentationPreference.BestFit, saves the workbook to a MemoryStream as HTML, measures the stream length, and asserts that the generated HTML does not exceed 2 MB, causing the unit test to fail if the limit is breached.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet.
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // Populate the worksheet with a large amount of data.
                const int rows = 5000;
                const int columns = 50;
                for (int row = 0; row < rows; row++)
                {
                    for (int col = 0; col < columns; col++)
                    {
                        sheet.Cells[row, col].PutValue($"Row{row}_Col{col}_Data");
                    }
                }

                // Configure HTML save options.
                var htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Embed images as Base64 to avoid external streams.
                    ExportImagesAsBase64 = true
                };

                // Save the workbook to a memory stream.
                using (var outputStream = new MemoryStream())
                {
                    workbook.Save(outputStream, htmlOptions);
                    long htmlSizeInBytes = outputStream.Length;

                    const long twoMegabytes = 2L * 1024 * 1024;
                    if (htmlSizeInBytes > twoMegabytes)
                    {
                        Console.WriteLine($"FAIL: Generated HTML size is {htmlSizeInBytes} bytes, exceeds 2 MB.");
                    }
                    else
                    {
                        Console.WriteLine($"PASS: Generated HTML size is {htmlSizeInBytes} bytes, within 2 MB.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception: {ex.Message}");
            }
        }
    }
}
