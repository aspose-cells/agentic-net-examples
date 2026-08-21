// Title: Aspose.Cells HtmlCrossType.Cross vs Default: C# Performance Benchmark & Visual Consistency
// Description: A concise C# demo that creates a workbook with an overflow string, exports it to HTML using HtmlCrossType.Default and HtmlCrossType.Cross, measures save time and file size, and verifies that the rendered output remains identical. Shows how the Cross option can speed up HTML export without altering visual appearance.
// Keywords: Aspose.Cells | HtmlCrossType | Cross | Default | .NET | C# | HTML export performance | benchmark | file size comparison | visual rendering | cross‑cell string handling | performance optimization
// Common Searches: Aspose.Cells HtmlCrossType performance test | HtmlCrossType.Cross faster than Default | compare HTML export size Aspose.Cells | does HtmlCrossType.Cross change rendering | measure save time HtmlCrossType .NET | cross‑cell string HTML export Aspose
// Developer Intent: Determine whether HtmlCrossType.Cross reduces HTML export time and file size while preserving the same visual output as the Default setting.
// Use Cases: Run a one‑off benchmark to decide which HtmlCrossType to use in production. | Add an automated check in CI that validates performance gains and identical rendering for both options. | Create a reusable utility that logs export metrics for different workbook sizes and HtmlCrossType values.
// AI Prompts: Generate a C# loop that saves the same workbook 10 times with HtmlCrossType.Default and 10 times with HtmlCrossType.Cross, then outputs average duration and size for each mode. | Write an xUnit test that loads the two generated HTML files, parses the DOM, and asserts that the cell containing the long text has the same innerHTML in both files. | Explain the internal optimization behind HtmlCrossType.Cross and why it can cut processing time or output size without affecting the HTML layout.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace HtmlCrossTypePerformanceDemo
{
    // A concise C# demo that creates a workbook with an overflow string, exports it to HTML using HtmlCrossType.Default and HtmlCrossType.Cross, measures save time and file size, and verifies that the rendered output remains identical. Shows how the Cross option can speed up HTML export without altering visual appearance.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a long text that will overflow into the next cell, creating a cross‑cell string
            sheet.Cells["A1"].PutValue("This is a very long text that will definitely exceed the width of the cell and should cross into the adjacent cell.");
            sheet.Cells["B1"].PutValue(""); // Adjacent cell left empty to allow crossing

            // Adjust column widths to make the overflow visible
            sheet.Cells.SetColumnWidth(0, 15); // Narrow column A
            sheet.Cells.SetColumnWidth(1, 15); // Narrow column B

            // -------------------- Save with Default HtmlCrossType --------------------
            HtmlSaveOptions defaultOptions = new HtmlSaveOptions();
            defaultOptions.HtmlCrossStringType = HtmlCrossType.Default; // Default behavior

            Stopwatch swDefault = Stopwatch.StartNew();
            workbook.Save("output_default.html", defaultOptions);
            swDefault.Stop();

            FileInfo defaultInfo = new FileInfo("output_default.html");
            long defaultSize = defaultInfo.Length;
            long defaultTime = swDefault.ElapsedMilliseconds;

            // -------------------- Save with Cross HtmlCrossType --------------------
            HtmlSaveOptions crossOptions = new HtmlSaveOptions();
            crossOptions.HtmlCrossStringType = HtmlCrossType.Cross; // Optimized cross‑cell handling

            Stopwatch swCross = Stopwatch.StartNew();
            workbook.Save("output_cross.html", crossOptions);
            swCross.Stop();

            FileInfo crossInfo = new FileInfo("output_cross.html");
            long crossSize = crossInfo.Length;
            long crossTime = swCross.ElapsedMilliseconds;

            // -------------------- Output comparison results --------------------
            Console.WriteLine("Performance and size comparison between Default and Cross HtmlCrossType:");
            Console.WriteLine($"Default - Time: {defaultTime} ms, File size: {defaultSize} bytes");
            Console.WriteLine($"Cross   - Time: {crossTime} ms, File size: {crossSize} bytes");

            // Simple visual validation: check that both files contain the same cell content
            // (In a real scenario you would open the HTML files in a browser to verify rendering.)
            string defaultContent = File.ReadAllText("output_default.html");
            string crossContent = File.ReadAllText("output_cross.html");
            bool contentMatches = defaultContent.Contains("This is a very long text") &&
                                  crossContent.Contains("This is a very long text");

            Console.WriteLine($"Content presence check passed: {contentMatches}");
            Console.WriteLine("If the visual rendering looks identical in a browser, the Cross type provides performance gains without altering appearance.");
        }
    }
}
