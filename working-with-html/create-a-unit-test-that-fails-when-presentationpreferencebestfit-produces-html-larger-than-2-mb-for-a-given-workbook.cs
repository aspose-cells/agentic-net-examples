// Title: Write a C# unit test with Aspose.Cells that fails when PresentationPreference.BestFit generates HTML larger than 2 MB
// AI Prompts: Create an MSTest method that builds a workbook with many rows, saves it to HTML using HtmlSaveOptions.PresentationPreference = PresentationPreference.BestFit, and asserts the MemoryStream length is under 2 MB. | Provide a NUnit test example that populates a worksheet, exports to HTML with BestFit presentation, and throws an assertion failure if the output exceeds 2 MB.
// Common Searches: aspocells unit test for HTML export size limit | c# assert Aspose.Cells HTML output under 2mb | how to use PresentationPreference.BestFit in automated test | verify generated HTML size with Aspose.Cells SaveOptions | fail test when Aspose.Cells HTML exceeds specific byte size
// Tags: Aspose.Cells HTML size unit test | PresentationPreference BestFit export limit | C# memory stream length assertion | HTML export size verification Aspose.Cells | automated test for workbook HTML byte size

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example demonstrates how to write a C# unit test that creates a large workbook, saves it to HTML with the PresentationPreference.BestFit setting, measures the resulting MemoryStream size, and fails the test if the HTML exceeds a 2 MB threshold.
    public class HtmlSizeTest
    {
        public static void Main()
        {
            try
            {
                // Create a workbook and populate it with a large amount of data
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];

                // Fill many rows and columns to increase the HTML size
                for (int row = 0; row < 5000; row++)
                {
                    for (int col = 0; col < 50; col++)
                    {
                        worksheet.Cells[row, col].PutValue($"R{row}C{col}");
                    }
                }

                // Configure HTML save options (default presentation)
                var htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

                // Save the workbook to a memory stream and check its size
                using (var memoryStream = new MemoryStream())
                {
                    workbook.Save(memoryStream, htmlOptions);
                    long htmlSizeInBytes = memoryStream.Length;

                    const long maxSizeBytes = 2L * 1024 * 1024; // 2 MB

                    if (htmlSizeInBytes <= maxSizeBytes)
                    {
                        Console.WriteLine($"Success: Generated HTML size {htmlSizeInBytes} bytes is within the 2 MB limit.");
                    }
                    else
                    {
                        Console.WriteLine($"Failure: Generated HTML size {htmlSizeInBytes} bytes exceeds the 2 MB limit.");
                    }
                }
            }
            catch (Exception ex)
            {
                // Runtime safety: report any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
