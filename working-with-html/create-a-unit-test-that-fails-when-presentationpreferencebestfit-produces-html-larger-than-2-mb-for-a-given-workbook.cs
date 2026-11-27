// Title: C# unit test that fails when Aspose.Cells PresentationPreference BestFit generates HTML larger than 2 MB
// Common Searches: how to create a C# unit test for Aspose.Cells PresentationPreference HTML size limit | verify Aspose.Cells HtmlSaveOptions output does not exceed 2 MB in a test | C# test to check Aspose.Cells BestFit HTML generation size using MemoryStream | unit testing Aspose.Cells HTML output size with PresentationPreference enabled | measure and assert Aspose.Cells generated HTML size in C# unit tests
// Tags: Aspose.Cells HtmlSaveOptions size validation | PresentationPreference HTML size limit | C# unit test Aspose.Cells HTML output | MemoryStream HTML size measurement Aspose.Cells | BestFit HTML generation size check

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPresentationPreferenceTest
{
    // The example builds a workbook with 5,000 rows and 50 columns, enables PresentationPreference in HtmlSaveOptions, saves the workbook to a MemoryStream as HTML, measures the stream length, and fails the unit test if the generated HTML exceeds the 2 MB threshold.
    public class HtmlSizeTest
    {
        // Size limit: 2 MB
        private const long MaxHtmlSizeBytes = 2L * 1024L * 1024L;

        public static void Run()
        {
            try
            {
                // 1. Create a workbook and populate it with a sizable data set.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                const int rows = 5000;
                const int cols = 50;

                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        sheet.Cells[r, c].PutValue($"R{r + 1}C{c + 1}");
                    }
                }

                // 2. Configure HtmlSaveOptions with PresentationPreference = true.
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    PresentationPreference = true
                };

                // 3. Save the workbook to a MemoryStream as HTML.
                using (MemoryStream htmlStream = new MemoryStream())
                {
                    workbook.Save(htmlStream, saveOptions);
                    long htmlSize = htmlStream.Length;

                    Console.WriteLine($"Generated HTML size: {htmlSize} bytes ({htmlSize / 1024.0 / 1024.0:F2} MB)");

                    // 4. Verify that the size does not exceed the 2 MB threshold.
                    if (htmlSize > MaxHtmlSizeBytes)
                    {
                        Console.WriteLine("FAIL: HTML size exceeds the allowed maximum of 2 MB.");
                        Environment.Exit(1);
                    }
                    else
                    {
                        Console.WriteLine("PASS: HTML size is within the allowed limit.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR: {ex.Message}");
                Environment.Exit(2);
            }
        }
    }

    public class Program
    {
        public static int Main(string[] args)
        {
            // Apply license if available (uncomment and set the correct path).
            // var license = new Aspose.Cells.License();
            // license.SetLicense("Aspose.Cells.lic");

            HtmlSizeTest.Run();
            return 0;
        }
    }
}
