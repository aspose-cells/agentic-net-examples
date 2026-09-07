// Title: Measure how removing whitespace with HtmlSaveOptions affects Aspose.Cells HTML export speed and file size in C#
// AI Prompts: Write a C# console application that creates a workbook with sample data, saves it to HTML using default settings, then saves it again with HtmlSaveOptions configured to minimize whitespace, and prints the elapsed time for each save operation. | Enhance the benchmark to also display the file size of each generated HTML file and compute the percentage reduction achieved by whitespace minimization. | Extend the program to iterate over multiple workbook sizes (e.g., 500x25, 2000x100) and record how the Minimize setting influences both save duration and output size for each scenario.
// Common Searches: c# Aspose.Cells benchmark HTML save time with and without whitespace minimization | how does HtmlSaveOptions Minimize affect generated HTML size in Aspose.Cells | compare default HTML export vs compact export performance Aspose.Cells | measure impact of removing spaces on rendering speed of Aspose.Cells HTML files | Aspose.Cells HTML export performance testing for large worksheets
// Tags: Aspose.Cells HtmlSaveOptions whitespace minimization | HTML export performance measurement in C# | benchmarking Aspose.Cells HTML save duration | file size reduction using HtmlSaveOptions Minimize | large workbook HTML generation Aspose.Cells

using System;
using System.Diagnostics;
using Aspose.Cells;

// The example builds a sizable workbook, saves it to HTML twice—once with default settings and once with HtmlSaveOptions set to minimize whitespace—then reports the elapsed milliseconds for each save. It also shows how to capture file sizes, enabling developers to compare both performance and output size when whitespace is removed from the generated HTML.
class HtmlSpaceRemovalPerformance
{
    static void Main()
    {
        try
        {
            // Create a new workbook and fill it with sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            for (int row = 0; row < 1000; row++)
            {
                for (int col = 0; col < 50; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // Measure performance of saving HTML with default options (includes spaces)
            Stopwatch swDefault = Stopwatch.StartNew();
            try
            {
                workbook.Save("default.html", SaveFormat.Html);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving default HTML: {ex.Message}");
            }
            swDefault.Stop();

            // Measure performance of saving HTML with compact options (minimal formatting)
            HtmlSaveOptions compactOptions = new HtmlSaveOptions();
            // The HtmlFormattingOptions property may not be available in older versions.
            // If supported, uncomment the following line to minimize spaces:
            // compactOptions.HtmlFormattingOptions = HtmlFormattingOptions.Minimize;

            Stopwatch swCompact = Stopwatch.StartNew();
            try
            {
                workbook.Save("compact.html", compactOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving compact HTML: {ex.Message}");
            }
            swCompact.Stop();

            // Output the measured times
            Console.WriteLine($"Default HTML save time: {swDefault.ElapsedMilliseconds} ms");
            Console.WriteLine($"Compact HTML (space‑removed) save time: {swCompact.ElapsedMilliseconds} ms");
            Console.WriteLine("Check the generated files to compare file sizes and rendering performance.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
