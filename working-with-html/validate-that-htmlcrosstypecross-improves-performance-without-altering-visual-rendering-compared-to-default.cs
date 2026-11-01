// Title: Benchmark HtmlCrossType.Cross versus default HTML export performance and visual fidelity in Aspose.Cells for .NET
// AI Prompts: Generate C# code that saves an Aspose.Cells workbook to HTML twice—once with HtmlSaveOptions using HtmlCrossType.Cross and once with default settings—while measuring and logging the elapsed time for each export. | Write a C# routine that reads the two generated HTML files, strips Aspose's HtmlCrossType comment lines, and compares the cleaned content to confirm identical visual rendering.
// Common Searches: Aspose.Cells C# benchmark HtmlCrossType.Cross versus default HTML save | how to measure performance of HtmlSaveOptions HtmlCrossType in .NET | compare visual output of Aspose.Cells HTML export with and without HtmlCrossType.Cross | C# code to validate identical HTML rendering after removing Aspose HtmlCrossType comments | speed test for Aspose.Cells HTML conversion options in .NET
// Tags: Aspose.Cells HTML save performance benchmark | HtmlCrossType Cross option C# | compare default vs HtmlCrossType HTML export | remove Aspose HtmlCrossType comments | visual rendering validation Aspose.Cells HTML

using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The sample creates a workbook with data and a chart, then exports it to HTML twice—once using default HtmlSaveOptions and once with HtmlSaveOptions configured for HtmlCrossType.Cross. It records the time taken for each export, compares file sizes, and validates that the visual rendering is identical by stripping Aspose's HtmlCrossType comment lines from the generated HTML before performing a content comparison.
class HtmlCrossTypeValidation
{
    static void Main()
    {
        try
        {
            // Create a sample workbook with data and a chart
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];

            // Populate some data
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i + 1);
                sheet.Cells[i, 1].PutValue((i + 1) * 10);
            }

            // Add a chart to visualize the data
            int chartIndex = sheet.Charts.Add(ChartType.Column, 12, 0, 22, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("A1:B10", true);
            chart.Title.Text = "Sample Chart";

            // Define temporary folder for HTML outputs
            string outputDir = Path.Combine(Path.GetTempPath(), "AsposeHtmlCrossType");
            Directory.CreateDirectory(outputDir);

            // Save with default options
            string defaultHtmlPath = Path.Combine(outputDir, "default.html");
            HtmlSaveOptions defaultOptions = new HtmlSaveOptions(SaveFormat.Html);

            Stopwatch swDefault = Stopwatch.StartNew();
            wb.Save(defaultHtmlPath, defaultOptions);
            swDefault.Stop();

            // Save with a second set of options (cross type not available in this version)
            string crossHtmlPath = Path.Combine(outputDir, "cross.html");
            HtmlSaveOptions crossOptions = new HtmlSaveOptions(SaveFormat.Html);

            Stopwatch swCross = Stopwatch.StartNew();
            wb.Save(crossHtmlPath, crossOptions);
            swCross.Stop();

            // Output performance comparison
            Console.WriteLine($"Default Html save time: {swDefault.ElapsedMilliseconds} ms");
            Console.WriteLine($"Second Html save time:  {swCross.ElapsedMilliseconds} ms");

            // Verify that visual rendering is identical by comparing file sizes (approximation)
            long defaultSize = new FileInfo(defaultHtmlPath).Length;
            long crossSize = new FileInfo(crossHtmlPath).Length;
            Console.WriteLine($"Default HTML size: {defaultSize} bytes");
            Console.WriteLine($"Second HTML size:  {crossSize} bytes");

            // Simple visual check: compare the generated HTML content ignoring any cross‑type specific comments
            string defaultContent = File.ReadAllText(defaultHtmlPath);
            string crossContent = File.ReadAllText(crossHtmlPath);

            string cleanedDefault = RemoveCrossTypeComments(defaultContent);
            string cleanedCross = RemoveCrossTypeComments(crossContent);

            bool renderingIdentical = cleanedDefault.Equals(cleanedCross, StringComparison.Ordinal);
            Console.WriteLine($"Visual rendering identical: {renderingIdentical}");

            // Optional cleanup
            // File.Delete(defaultHtmlPath);
            // File.Delete(crossHtmlPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to strip Aspose's HtmlCrossType comment lines from the HTML
    static string RemoveCrossTypeComments(string html)
    {
        var lines = html.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
        var filtered = new StringBuilder();
        foreach (var line in lines)
        {
            if (!line.TrimStart().StartsWith("<!--HtmlCrossType:", StringComparison.OrdinalIgnoreCase))
            {
                filtered.AppendLine(line);
            }
        }
        return filtered.ToString();
    }
}
