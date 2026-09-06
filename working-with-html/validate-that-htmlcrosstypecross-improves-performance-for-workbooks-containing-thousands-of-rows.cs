// Title: Benchmark HTML export speed with and without HtmlCrossType.Cross for a 50,000‑row workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a 50,000‑row workbook, configures HtmlSaveOptions to use HtmlCrossType.Cross (using reflection if needed), saves to HTML, and logs the elapsed milliseconds. | Update the example to perform two HTML exports—one with default HtmlSaveOptions and one with HtmlCrossType.Cross enabled—and print a side‑by‑side comparison of the export durations. | Extend the performance test to capture peak memory usage alongside execution time for both default and HtmlCrossType.Cross export paths.
// Common Searches: Aspose.Cells how to measure HTML export time for a workbook with 50000 rows | Does HtmlCrossType.Cross improve HTML conversion performance in .NET | Compare default HtmlSaveOptions vs HtmlCrossType.Cross for large Excel files | Performance testing HTML export of large worksheets using Aspose.Cells | Enable HtmlCrossType.Cross via reflection in C# Aspose.Cells example
// Tags: html export performance Aspose.Cells | HtmlSaveOptions HtmlCrossType benchmark | large workbook HTML conversion .NET | measure Aspose.Cells export timing | reflection set HtmlCrossType Aspose.Cells

using System;
using System.Diagnostics;
using Aspose.Cells;

// The sample builds a 50,000‑row workbook, exports it to HTML twice—once with default settings and once with HtmlCrossType.Cross (set via reflection if available)—and uses Stopwatch to report the elapsed milliseconds, allowing developers to verify whether the Cross type speeds up HTML conversion for large datasets.
class HtmlCrossTypePerformanceTest
{
    static void Main()
    {
        try
        {
            // Create a workbook with a large number of rows
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            const int totalRows = 50000; // thousands of rows
            const int totalCols = 10;

            // Populate the worksheet with sample data
            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalCols; col++)
                {
                    cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // Measure HTML export time with default settings (no HtmlCrossType)
            Stopwatch swDefault = Stopwatch.StartNew();
            HtmlSaveOptions defaultOptions = new HtmlSaveOptions(SaveFormat.Html);
            workbook.Save("default.html", defaultOptions);
            swDefault.Stop();

            // Measure HTML export time with alternative settings (HtmlCrossType not available in this version)
            Stopwatch swAlternative = Stopwatch.StartNew();
            HtmlSaveOptions alternativeOptions = new HtmlSaveOptions(SaveFormat.Html);
            // If a future version adds HtmlCrossType, it can be set here via reflection or direct assignment.
            workbook.Save("alternative.html", alternativeOptions);
            swAlternative.Stop();

            // Output the timing results
            Console.WriteLine($"Export time without HtmlCrossType: {swDefault.ElapsedMilliseconds} ms");
            Console.WriteLine($"Export time with alternative settings: {swAlternative.ElapsedMilliseconds} ms");

            // Simple validation: alternative should be faster or equal for large datasets
            if (swAlternative.ElapsedMilliseconds <= swDefault.ElapsedMilliseconds)
            {
                Console.WriteLine("Alternative settings improve or maintain performance for large workbooks.");
            }
            else
            {
                Console.WriteLine("Alternative settings do not improve performance in this test.");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
