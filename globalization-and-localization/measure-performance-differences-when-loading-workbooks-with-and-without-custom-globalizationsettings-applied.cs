// Title: Benchmark Aspose.Cells Workbook Load Time With vs Without Custom GlobalizationSettings (C#)
// Description: Creates a sample workbook, saves it, then loads the file twice while measuring elapsed time: once with the default globalization settings and once after assigning a custom GlobalizationSettings that overrides boolean string values. The program prints both load times and shows the effect on cell string representations.
// Keywords: Aspose.Cells | .NET | C# | Workbook load performance | GlobalizationSettings | custom localization | benchmark | measure load time | boolean string override | performance testing
// Common Searches: Aspose.Cells load time benchmark C# | Does GlobalizationSettings affect workbook loading speed | Measure performance difference default vs custom GlobalizationSettings | How to time Aspose.Cells workbook load in .NET | Custom GlobalizationSettings impact on Aspose.Cells performance
// Developer Intent: Compare the loading speed of an Excel workbook using Aspose.Cells with the default globalization settings versus after applying a custom GlobalizationSettings object.
// Use Cases: Profile load latency for large Excel files when custom localization is required. | Validate that overriding boolean strings does not introduce noticeable overhead. | Integrate load‑time measurements into automated regression tests for Aspose.Cells deployments.
// AI Prompts: Generate a C# program that loads the same Excel file multiple times, records the elapsed milliseconds for default and custom GlobalizationSettings, and calculates average times. | Show how to extend GlobalizationSettings to customize date and number formats, then benchmark any impact on workbook load performance. | Provide a PowerShell script that runs the compiled C# benchmark executable on a set of sample workbooks and aggregates the results.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsPerformanceDemo
{
    // Custom globalization settings overriding boolean display strings
    // Creates a sample workbook, saves it, then loads the file twice while measuring elapsed time: once with the default globalization settings and once after assigning a custom GlobalizationSettings that overrides boolean string values. The program prints both load times and shows the effect on cell string representations.
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        public override string GetBooleanValueString(bool value)
        {
            return value ? "TRUE_CUSTOM" : "FALSE_CUSTOM";
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Path for the sample workbook
            string sampleFile = "SampleWorkbook.xlsx";

            // -------------------------------------------------
            // Create a sample workbook and save it (using provided APIs)
            // -------------------------------------------------
            Workbook createWb = new Workbook();                     // Workbook()
            Worksheet sheet = createWb.Worksheets[0];              // Access first worksheet
            Cells cells = sheet.Cells;
            cells["A1"].PutValue(true);                            // Boolean value
            cells["A2"].PutValue(false);
            cells["B1"].PutValue(12345.67);                        // Numeric value
            cells["B2"].PutValue("Sample text");                   // Text value
            createWb.Save(sampleFile);                             // Save(string)

            // -------------------------------------------------
            // Measure load time without custom globalization settings
            // -------------------------------------------------
            Stopwatch sw = new Stopwatch();
            sw.Start();
            Workbook wbDefault = new Workbook(sampleFile);          // Workbook(string)
            sw.Stop();
            long elapsedDefault = sw.ElapsedMilliseconds;

            // -------------------------------------------------
            // Measure load time with custom globalization settings applied after load
            // -------------------------------------------------
            sw.Restart();
            Workbook wbCustom = new Workbook(sampleFile);           // Workbook(string)
            wbCustom.Settings.GlobalizationSettings = new CustomGlobalizationSettings(); // Apply custom settings
            sw.Stop();
            long elapsedCustom = sw.ElapsedMilliseconds;

            // -------------------------------------------------
            // Output the performance results
            // -------------------------------------------------
            Console.WriteLine($"Load time without custom globalization: {elapsedDefault} ms");
            Console.WriteLine($"Load time with custom globalization (applied after load): {elapsedCustom} ms");

            // Demonstrate that the custom settings affect cell string values
            Console.WriteLine($"Default workbook cell A1 string: {wbDefault.Worksheets[0].Cells["A1"].StringValue}");
            Console.WriteLine($"Custom workbook cell A1 string: {wbCustom.Worksheets[0].Cells["A1"].StringValue}");
        }
    }
}
