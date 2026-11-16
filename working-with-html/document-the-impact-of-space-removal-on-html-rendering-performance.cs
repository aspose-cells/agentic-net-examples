// Title: Measure the effect of whitespace removal on HTML rendering performance for Aspose.Cells output using C#
// Common Searches: how to benchmark HTML load time after minifying Aspose.Cells generated HTML in C# | C# code to compare original and minified HTML file sizes from an Aspose.Cells workbook | measure impact of whitespace reduction on HTML rendering speed with .NET | using Regex to minify Aspose.Cells HTML and evaluate performance | performance testing of reading Aspose.Cells HTML files before and after whitespace cleanup
// Tags: Aspose.Cells HTML minification C# | measure HTML read latency .NET | regex whitespace collapse for HTML | compare original and minified HTML size | benchmark Aspose.Cells HTML output performance

using System;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example creates a workbook, saves it as HTML, times the file read operation, minifies the HTML by collapsing whitespace with a regular expression, saves the minified file, and reports both original and minified file sizes along with the percentage reduction, illustrating the performance impact of whitespace removal.
class Program
{
    static void Main()
    {
        // Create a workbook and populate it with sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        for (int row = 0; row < 1000; row++)
        {
            for (int col = 0; col < 10; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row}C{col}");
            }
        }

        // Save the workbook as HTML (original version with default formatting)
        string originalHtmlPath = "original.html";
        workbook.Save(originalHtmlPath, SaveFormat.Html);

        // Measure the time required to read the original HTML file (simulating rendering load time)
        Stopwatch stopwatch = new Stopwatch();
        stopwatch.Start();
        string originalHtml = File.ReadAllText(originalHtmlPath);
        stopwatch.Stop();
        Console.WriteLine($"Read original HTML time: {stopwatch.ElapsedMilliseconds} ms");

        // Remove unnecessary whitespace to reduce HTML size
        stopwatch.Restart();
        string minifiedHtml = Regex.Replace(originalHtml, @"\s+", " ");
        stopwatch.Stop();
        Console.WriteLine($"HTML minification time: {stopwatch.ElapsedMilliseconds} ms");

        // Save the minified HTML
        string minifiedHtmlPath = "minified.html";
        File.WriteAllText(minifiedHtmlPath, minifiedHtml);

        // Compare file sizes to illustrate performance impact
        long originalSize = new FileInfo(originalHtmlPath).Length;
        long minifiedSize = new FileInfo(minifiedHtmlPath).Length;
        Console.WriteLine($"Original HTML size: {originalSize} bytes");
        Console.WriteLine($"Minified HTML size: {minifiedSize} bytes");
        Console.WriteLine($"Size reduction: {originalSize - minifiedSize} bytes ({(originalSize - minifiedSize) * 100.0 / originalSize:0.##}%)");
    }
}
