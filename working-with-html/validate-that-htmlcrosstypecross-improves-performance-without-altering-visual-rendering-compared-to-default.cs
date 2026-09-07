// Title: Measure performance impact of HtmlCrossType.Cross vs Default when exporting Excel to HTML using Aspose.Cells for .NET
// AI Prompts: Generate a C# console program that loads an Excel workbook, saves it to HTML twice with Aspose.Cells—once using HtmlCrossType.Default and once using HtmlCrossType.Cross—records the elapsed time and output file size for each, and prints a summary indicating which option is faster. | Write C# code that benchmarks Aspose.Cells HTML export by timing Workbook.Save with HtmlCrossType.Cross and HtmlCrossType.Default, compares the generated HTML file sizes, and logs the results to the console.
// Common Searches: Aspose.Cells how to compare HtmlCrossType.Cross and Default export speed in C# | C# benchmark Excel to HTML conversion using HtmlCrossType.Cross | measure Aspose.Cells HTML export performance HtmlCrossType setting | does HtmlCrossType.Cross reduce save time for large workbooks in .NET | compare file size of HTML output with HtmlCrossType.Cross vs Default Aspose.Cells
// Tags: Aspose.Cells HtmlCrossType performance benchmark | C# Excel to HTML export timing | HtmlCrossType.Cross vs Default speed test | Aspose.Cells HTML file size comparison | measure workbook.save latency Aspose.Cells

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The program loads Sample.xlsx, exports it to HTML twice using Aspose.Cells with HtmlCrossType set to Default and Cross (via reflection), measures each save operation with Stopwatch, records the resulting file sizes, and outputs which setting provides a faster export without changing visual rendering.
class HtmlCrossTypePerformanceTest
{
    static void Main()
    {
        // Path to the source workbook
        string workbookPath = "Sample.xlsx";

        // Verify that the workbook file exists
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: Workbook file \"{workbookPath}\" not found.");
            return;
        }

        Workbook workbook;
        try
        {
            workbook = new Workbook(workbookPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Prepare HTML save options for the default cross type
        HtmlSaveOptions defaultOptions = new HtmlSaveOptions(SaveFormat.Html);
        defaultOptions.ExportActiveWorksheetOnly = true;
        // Set HtmlCrossType to Default via reflection (if the property exists in the current version)
        var htmlCrossProp = typeof(HtmlSaveOptions).GetProperty("HtmlCrossType");
        if (htmlCrossProp != null)
        {
            htmlCrossProp.SetValue(defaultOptions, HtmlCrossType.Default);
        }

        // Prepare HTML save options for the Cross cross type
        HtmlSaveOptions crossOptions = new HtmlSaveOptions(SaveFormat.Html);
        crossOptions.ExportActiveWorksheetOnly = true;
        // Set HtmlCrossType to Cross via reflection (if supported)
        if (htmlCrossProp != null)
        {
            htmlCrossProp.SetValue(crossOptions, HtmlCrossType.Cross);
        }

        // Measure performance for Default
        Stopwatch swDefault = Stopwatch.StartNew();
        try
        {
            workbook.Save("Default.html", defaultOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving Default HTML: {ex.Message}");
            return;
        }
        swDefault.Stop();

        // Measure performance for Cross
        Stopwatch swCross = Stopwatch.StartNew();
        try
        {
            workbook.Save("Cross.html", crossOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving Cross HTML: {ex.Message}");
            return;
        }
        swCross.Stop();

        // Output timing results
        Console.WriteLine($"Default HtmlCrossType save time: {swDefault.ElapsedMilliseconds} ms");
        Console.WriteLine($"Cross HtmlCrossType save time: {swCross.ElapsedMilliseconds} ms");

        // Compare file sizes as a simple proxy for rendering output size
        long sizeDefault = new FileInfo("Default.html").Length;
        long sizeCross = new FileInfo("Cross.html").Length;
        Console.WriteLine($"Default HTML file size: {sizeDefault} bytes");
        Console.WriteLine($"Cross HTML file size: {sizeCross} bytes");

        // Final assessment based on performance (visual comparison omitted)
        if (swCross.ElapsedMilliseconds < swDefault.ElapsedMilliseconds)
        {
            Console.WriteLine("HtmlCrossType.Cross improves performance.");
        }
        else
        {
            Console.WriteLine("HtmlCrossType.Cross does not provide a clear performance benefit.");
        }
    }
}
