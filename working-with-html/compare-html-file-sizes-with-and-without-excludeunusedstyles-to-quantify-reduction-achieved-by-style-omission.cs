// Title: Calculate HTML size reduction from ExcludeUnusedStyles when converting Excel to HTML with Aspose.Cells for .NET (C#)
// Common Searches: C# how to compare HTML output size with and without ExcludeUnusedStyles in Aspose.Cells | Aspose.Cells measure HTML file size reduction after enabling ExcludeUnusedStyles | calculate percentage decrease of generated HTML using HtmlSaveOptions.ExcludeUnusedStyles | compare default and optimized HTML files from Excel conversion in .NET | determine bytes saved by excluding unused styles in Aspose.Cells HTML export
// Tags: Aspose.Cells HtmlSaveOptions ExcludeUnusedStyles | HTML output size optimization Aspose.Cells | C# compare generated HTML file sizes | measure unused style omission impact | Excel to HTML conversion size reduction

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, saves it twice as HTML—once with ExcludeUnusedStyles disabled (default) and once with it enabled—then reads the byte length of each file, calculates the absolute and percentage size reduction, and prints the results to the console.
class HtmlStyleComparison
{
    static void Main()
    {
        // Path to the source Excel file
        string excelPath = "input.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(excelPath);

        // Directory to store generated HTML files
        string outputDir = "HtmlOutputs";
        Directory.CreateDirectory(outputDir);

        // Save HTML without excluding unused styles (default behavior)
        string htmlPathDefault = Path.Combine(outputDir, "default.html");
        HtmlSaveOptions optionsDefault = new HtmlSaveOptions();
        // Ensure default setting (ExcludeUnusedStyles = false)
        optionsDefault.ExcludeUnusedStyles = false;
        workbook.Save(htmlPathDefault, optionsDefault);

        // Save HTML with ExcludeUnusedStyles enabled
        string htmlPathOptimized = Path.Combine(outputDir, "optimized.html");
        HtmlSaveOptions optionsOptimized = new HtmlSaveOptions();
        optionsOptimized.ExcludeUnusedStyles = true;
        workbook.Save(htmlPathOptimized, optionsOptimized);

        // Get file sizes
        long sizeDefault = new FileInfo(htmlPathDefault).Length;
        long sizeOptimized = new FileInfo(htmlPathOptimized).Length;

        // Calculate reduction
        long reductionBytes = sizeDefault - sizeOptimized;
        double reductionPercent = sizeDefault > 0
            ? (double)reductionBytes / sizeDefault * 100
            : 0;

        // Output results
        Console.WriteLine($"HTML size without ExcludeUnusedStyles: {sizeDefault} bytes");
        Console.WriteLine($"HTML size with ExcludeUnusedStyles:    {sizeOptimized} bytes");
        Console.WriteLine($"Size reduction: {reductionBytes} bytes ({reductionPercent:F2}%)");
    }
}
