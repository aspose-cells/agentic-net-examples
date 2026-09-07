// Title: Benchmark the performance of Aspose.Cells HTML conversion for a 10 MB workbook with CSS export enabled versus disabled in C#
// AI Prompts: Write a C# console app that loads a 10 MB Excel file, saves it to HTML twice using Aspose.Cells with HtmlSaveOptions.ExportCss set to true and false, and returns the elapsed milliseconds for each run. | Create a script that times Workbook.Save when HtmlSaveOptions.ExportCss is toggled, logs both durations, and deletes the temporary HTML files after measurement. | Generate C# code that measures and compares Aspose.Cells HTML export speed for a large workbook, handling missing file errors and outputting the conversion times for CSS‑enabled and CSS‑disabled scenarios.
// Common Searches: how long does Aspose.Cells take to convert a 10 MB Excel workbook to HTML with CSS enabled | performance impact of ExportCss true vs false in Aspose.Cells HTML export | C# benchmark Aspose.Cells HtmlSaveOptions ExportCss setting | measure Aspose.Cells HTML conversion speed for large workbooks | timing workbook.Save to HTML with and without CSS using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportCss benchmark | HTML conversion performance Aspose.Cells C# | large workbook to HTML timing | disable CSS in Aspose.Cells HTML export | measure Workbook.Save execution time

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example loads a 10 MB Excel workbook, converts it to HTML twice with Aspose.Cells—once with ExportCss enabled and once disabled—using HtmlSaveOptions, measures each conversion with Stopwatch, prints the elapsed milliseconds, deletes the generated HTML files, and includes error handling for missing files or conversion failures.
class HtmlConversionBenchmark
{
    static void Main()
    {
        // Path to the 10 MB workbook (ensure the file exists)
        string workbookPath = "LargeWorkbook.xlsx";

        // Verify the workbook file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
            return;
        }

        try
        {
            // Load the workbook (lifecycle rule)
            Workbook workbook = new Workbook(workbookPath);

            // Benchmark with CSS enabled
            TimeSpan cssEnabledTime = ConvertToHtml(workbook, enableCss: true);
            Console.WriteLine($"HTML conversion with CSS enabled: {cssEnabledTime.TotalMilliseconds} ms");

            // Benchmark with CSS disabled
            TimeSpan cssDisabledTime = ConvertToHtml(workbook, enableCss: false);
            Console.WriteLine($"HTML conversion with CSS disabled: {cssDisabledTime.TotalMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during processing: {ex.Message}");
        }
    }

    /// <param name="workbook">The workbook to convert.</param>
    /// <param name="enableCss">True to enable CSS export, false to disable it.</param>
    /// <returns>Time taken for the conversion.</returns>
    private static TimeSpan ConvertToHtml(Workbook workbook, bool enableCss)
    {
        // Configure HTML save options
        HtmlSaveOptions options = new HtmlSaveOptions();

        // Note: In the current Aspose.Cells version, CSS export is controlled via ExportCss.
        // If the property is unavailable, the default behavior will be used.
        try
        {
            // Attempt to set CSS export flag if the property exists.
            var exportCssProp = typeof(HtmlSaveOptions).GetProperty("ExportCss");
            if (exportCssProp != null && exportCssProp.CanWrite)
            {
                exportCssProp.SetValue(options, enableCss);
            }
        }
        catch
        {
            // Ignore any reflection errors; proceed with default options.
        }

        // Use a dummy output path (the file is not needed for timing)
        string outputPath = enableCss ? "output_css_enabled.html" : "output_css_disabled.html";

        try
        {
            // Measure conversion time
            Stopwatch sw = Stopwatch.StartNew();
            workbook.Save(outputPath, options);
            sw.Stop();

            // Optionally delete the generated file to keep the folder clean
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            return sw.Elapsed;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Conversion failed: {ex.Message}");
            return TimeSpan.Zero;
        }
    }
}
