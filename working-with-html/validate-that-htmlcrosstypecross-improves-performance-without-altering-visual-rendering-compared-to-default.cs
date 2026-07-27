// Title: Benchmark HtmlCrossStringType.Cross vs Default for Aspose.Cells HTML Export (C#)
// Description: A C# console app that creates a workbook with overflow text, saves it to HTML using HtmlCrossStringType.Default and HtmlCrossStringType.Cross, measures execution time and file size, and verifies that the visual output remains identical.
// Keywords: Aspose.Cells | C# | HtmlCrossStringType | Cross vs Default | HTML export performance | benchmark | visual rendering consistency
// Common Searches: Aspose.Cells HtmlCrossStringType performance test | Does HtmlCrossStringType.Cross change HTML appearance | Measure HTML export speed Aspose.Cells C# | Compare file size for HtmlCrossStringType.Cross and Default | How to validate visual output of Aspose.Cells HTML export
// Developer Intent: Find out whether HtmlCrossStringType.Cross speeds up HTML generation without affecting the rendered result compared to the Default setting.
// Use Cases: Run a quick benchmark to choose the optimal HtmlCrossStringType for large worksheets. | Integrate visual‑equivalence checks into CI pipelines for Aspose.Cells HTML exports. | Collect size and timing metrics to guide performance‑focused deployment decisions.
// AI Prompts: Generate a C# method that compares two HTML files produced by Aspose.Cells, ignoring any internal markers related to HtmlCrossStringType. | Write a PowerShell script that executes the benchmark program multiple times and reports average export time and file size for both HtmlCrossStringType options. | Explain how HtmlCrossStringType.Cross reduces processing overhead when handling overflow text during HTML conversion.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlCrossTypeValidation
{
    // A C# console app that creates a workbook with overflow text, saves it to HTML using HtmlCrossStringType.Default and HtmlCrossStringType.Cross, measures execution time and file size, and verifies that the visual output remains identical.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data that will cause cross‑cell strings
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Long text in A1 that will overflow into B1
            sheet.Cells["A1"].PutValue("This is a very long text that will cross into the next cell when rendered as HTML.");
            sheet.Cells["B1"].PutValue("Adjacent cell");

            // Apply a thin border to visualize the cells in HTML
            Style borderStyle = workbook.CreateStyle();
            borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            sheet.Cells["A1"].SetStyle(borderStyle);
            sheet.Cells["B1"].SetStyle(borderStyle);

            // Prepare HTML save options for Default cross type
            HtmlSaveOptions defaultOptions = new HtmlSaveOptions();
            defaultOptions.HtmlCrossStringType = HtmlCrossType.Default;

            // Measure time and size for Default
            Stopwatch sw = new Stopwatch();
            sw.Start();
            string defaultFile = "output_default.html";
            workbook.Save(defaultFile, defaultOptions);
            sw.Stop();
            long defaultTimeMs = sw.ElapsedMilliseconds;
            long defaultSize = new FileInfo(defaultFile).Length;

            // Prepare HTML save options for Cross cross type
            HtmlSaveOptions crossOptions = new HtmlSaveOptions();
            crossOptions.HtmlCrossStringType = HtmlCrossType.Cross;

            // Measure time and size for Cross
            sw.Restart();
            string crossFile = "output_cross.html";
            workbook.Save(crossFile, crossOptions);
            sw.Stop();
            long crossTimeMs = sw.ElapsedMilliseconds;
            long crossSize = new FileInfo(crossFile).Length;

            // Output the performance comparison
            Console.WriteLine($"Default  - Time: {defaultTimeMs} ms, Size: {defaultSize} bytes");
            Console.WriteLine($"Cross    - Time: {crossTimeMs} ms, Size: {crossSize} bytes");

            // Simple visual validation: compare the two HTML files line by line ignoring known differences
            // (e.g., the cross‑type attribute does not affect visible HTML markup)
            bool visualMatch = FilesAreVisuallyEquivalent(defaultFile, crossFile);
            Console.WriteLine($"Visual rendering unchanged: {visualMatch}");
        }

        // Helper method to compare two HTML files while ignoring the HtmlCrossStringType attribute differences
        static bool FilesAreVisuallyEquivalent(string file1, string file2)
        {
            string[] lines1 = File.ReadAllLines(file1);
            string[] lines2 = File.ReadAllLines(file2);

            if (lines1.Length != lines2.Length)
                return false;

            for (int i = 0; i < lines1.Length; i++)
            {
                // Remove any attribute that may contain the cross‑type setting
                string cleaned1 = RemoveCrossTypeAttribute(lines1[i]);
                string cleaned2 = RemoveCrossTypeAttribute(lines2[i]);

                if (!string.Equals(cleaned1, cleaned2, StringComparison.Ordinal))
                    return false;
            }
            return true;
        }

        // Strips possible HtmlCrossStringType related markers (if any) from a line
        static string RemoveCrossTypeAttribute(string line)
        {
            // The HtmlCrossStringType does not embed explicit markers in the HTML,
            // but this method is kept for completeness in case future versions add them.
            return line;
        }
    }
}
