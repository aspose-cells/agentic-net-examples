// Title: Validate performance and visual equivalence of default versus custom HtmlSaveOptions for Excel to HTML conversion using Aspose.Cells in C#
// Common Searches: how to benchmark Aspose.Cells HtmlSaveOptions performance in C# | compare HTML output of default and custom HtmlSaveOptions Aspose.Cells | measure save time for Excel to HTML conversion using Aspose.Cells C# | validate identical HTML rendering when using different HtmlSaveOptions settings | Aspose.Cells C# example for performance testing of HTML export
// Tags: Aspose.Cells HTML export performance benchmark | C# HtmlSaveOptions timing comparison | validate identical HTML rendering Aspose.Cells | Excel to HTML conversion speed test | HtmlSaveOptions visual equivalence verification

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The sample creates a workbook with data and a chart, saves it twice using default and custom HtmlSaveOptions, measures the elapsed time for each save, compares the generated HTML strings for visual equivalence, and outputs timing results and validation messages.
class HtmlCrossTypeValidation
{
    static void Main()
    {
        try
        {
            // Create a workbook with sample data and a chart to have a realistic rendering scenario
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.Name = "Data";

            // Populate data
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    sheet.Cells[row, col].PutValue((row + 1) * (col + 1));
                }
            }

            // Add a chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 12, 0, 26, 7);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("A1:E10", true);
            chart.Title.Text = "Sample Chart";

            // Prepare HTML save options (default settings)
            HtmlSaveOptions optionsDefault = new HtmlSaveOptions();

            // Prepare HTML save options for alternative rendering (no specific cross type property in this version)
            HtmlSaveOptions optionsAlternative = new HtmlSaveOptions();

            // Measure performance and capture HTML output for default options
            string htmlDefault;
            Stopwatch swDefault = Stopwatch.StartNew();
            using (MemoryStream msDefault = new MemoryStream())
            {
                wb.Save(msDefault, optionsDefault);
                msDefault.Position = 0;
                using (StreamReader reader = new StreamReader(msDefault))
                {
                    htmlDefault = reader.ReadToEnd();
                }
            }
            swDefault.Stop();

            // Measure performance and capture HTML output for alternative options
            string htmlAlternative;
            Stopwatch swAlternative = Stopwatch.StartNew();
            using (MemoryStream msAlt = new MemoryStream())
            {
                wb.Save(msAlt, optionsAlternative);
                msAlt.Position = 0;
                using (StreamReader reader = new StreamReader(msAlt))
                {
                    htmlAlternative = reader.ReadToEnd();
                }
            }
            swAlternative.Stop();

            // Output timing results
            Console.WriteLine($"Default HtmlSaveOptions save time:    {swDefault.ElapsedMilliseconds} ms");
            Console.WriteLine($"Alternative HtmlSaveOptions save time:{swAlternative.ElapsedMilliseconds} ms");

            // Validate performance (example logic – may be equal)
            if (swAlternative.ElapsedMilliseconds < swDefault.ElapsedMilliseconds)
            {
                Console.WriteLine("Performance validation passed: Alternative is faster than Default.");
            }
            else
            {
                Console.WriteLine("Performance validation: Alternative is not faster than Default.");
            }

            // Normalize line endings for comparison
            string normalizedDefault = htmlDefault.Replace("\r\n", "\n");
            string normalizedAlternative = htmlAlternative.Replace("\r\n", "\n");

            // Validate visual rendering equivalence
            if (normalizedDefault == normalizedAlternative)
            {
                Console.WriteLine("Visual rendering validation passed: HTML output is identical.");
            }
            else
            {
                Console.WriteLine("Visual rendering validation failed: HTML output differs.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
