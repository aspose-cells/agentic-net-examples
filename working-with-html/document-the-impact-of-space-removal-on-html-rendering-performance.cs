// Title: Compare Aspose.Cells HTML export speed and output size before and after whitespace removal using C#
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, saves it to HTML, then creates a second HTML version by stripping redundant spaces with a regular expression, and records the elapsed time for each step. | Write a C# snippet that prints both the generation time and the file size of the original HTML and the whitespace‑collapsed HTML produced by Aspose.Cells.
// Common Searches: how to benchmark Aspose.Cells HTML conversion time in C# | C# compare size of Aspose.Cells generated HTML with and without minification | measure performance impact of removing spaces from HTML saved by Aspose.Cells | regex to collapse whitespace between HTML tags after exporting Excel with Aspose.Cells | Aspose.Cells HTML export file size difference after whitespace cleanup
// Tags: Aspose.Cells HTML export performance | C# whitespace removal from generated HTML | HTML file size comparison after Aspose.Cells conversion | benchmark Excel to HTML conversion timing | C# regex for HTML whitespace reduction

using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, saves it to HTML twice, measures the time taken for each save, removes redundant whitespace from the second output using regular expressions, writes both versions to disk, and displays the generation times and file sizes to illustrate the impact of whitespace removal on performance and output size.
class HtmlSpaceRemovalPerformance
{
    static void Main()
    {
        // Path to the source Excel file
        string excelPath = "Sample.xlsx";

        // Load the workbook (lifecycle rule: load)
        Workbook workbook = new Workbook(excelPath);

        // Prepare HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Export the entire workbook to a single HTML page
            ExportActiveWorksheetOnly = false,
            // Keep the original formatting
            ExportImagesAsBase64 = true
        };

        // Measure time to save with default formatting (including spaces)
        Stopwatch swDefault = Stopwatch.StartNew();
        using (MemoryStream msDefault = new MemoryStream())
        {
            workbook.Save(msDefault, htmlOptions);
            swDefault.Stop();

            // Get the generated HTML as a string
            string htmlDefault = Encoding.UTF8.GetString(msDefault.ToArray());

            // Write the default HTML to a file for reference
            File.WriteAllText("Output_Default.html", htmlDefault, Encoding.UTF8);
        }

        // Measure time to save and then remove unnecessary spaces
        Stopwatch swTrimmed = Stopwatch.StartNew();
        using (MemoryStream msTrimmed = new MemoryStream())
        {
            workbook.Save(msTrimmed, htmlOptions);
            swTrimmed.Stop();

            // Convert to string
            string htmlTrimmed = Encoding.UTF8.GetString(msTrimmed.ToArray());

            // Remove redundant whitespace (spaces, tabs, line breaks) between tags
            // This regex collapses multiple whitespace characters into a single space
            // and removes spaces between '>' and '<' to tighten the markup.
            htmlTrimmed = Regex.Replace(htmlTrimmed, @">\s+<", "><");
            htmlTrimmed = Regex.Replace(htmlTrimmed, @"\s{2,}", " ");

            // Write the trimmed HTML to a file
            File.WriteAllText("Output_Trimmed.html", htmlTrimmed, Encoding.UTF8);
        }

        // Output the measured times
        Console.WriteLine($"HTML generation time (with spaces): {swDefault.ElapsedMilliseconds} ms");
        Console.WriteLine($"HTML generation time (spaces removed): {swTrimmed.ElapsedMilliseconds} ms");
        Console.WriteLine("Check Output_Default.html and Output_Trimmed.html for size differences.");

        // Optionally, display file size comparison
        FileInfo defaultInfo = new FileInfo("Output_Default.html");
        FileInfo trimmedInfo = new FileInfo("Output_Trimmed.html");
        Console.WriteLine($"Default HTML size: {defaultInfo.Length} bytes");
        Console.WriteLine($"Trimmed HTML size: {trimmedInfo.Length} bytes");
    }
}
