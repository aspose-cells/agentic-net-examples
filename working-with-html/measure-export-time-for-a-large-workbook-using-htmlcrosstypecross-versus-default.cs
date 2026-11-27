// Title: Benchmark HTML export time for a large workbook using HtmlCrossType.Cross vs HtmlCrossType.Default in Aspose.Cells for .NET
// Common Searches: Aspose.Cells compare HtmlCrossType.Cross and HtmlCrossType.Default export speed in C# | measure HTML save performance for a large workbook with Aspose.Cells | how long does HtmlCrossType.Cross take to export a 5000x50 sheet in .NET | benchmark Aspose.Cells HTML export using Stopwatch | set HtmlCrossType property via reflection in Aspose.Cells example
// Tags: Aspose.Cells HTML export performance | HtmlCrossType.Cross benchmark | large workbook HTML save Aspose.Cells | C# Stopwatch timing Aspose.Cells export | reflection set HtmlCrossType property

using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Rendering;   // Required for HtmlCrossType enum

// The sample creates a 5000‑row by 50‑column workbook, fills it with data, then saves it twice to HTML—once with HtmlCrossType.Default and once with HtmlCrossType.Cross—using reflection to set the HtmlCrossType property when available, and records the duration of each export with a Stopwatch.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define size of the large workbook
            const int totalRows = 5000;   // adjust as needed for "large"
            const int totalCols = 50;     // adjust as needed for "large"

            // Populate the worksheet with sample data
            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalCols; col++)
                {
                    worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // -----------------------------------------------------------------
            // Export using HtmlCrossType.Default (explicitly set if property exists)
            // -----------------------------------------------------------------
            HtmlSaveOptions defaultOptions = new HtmlSaveOptions(SaveFormat.Html);
            SetHtmlCrossTypeIfSupported(defaultOptions, HtmlCrossType.Default);

            Stopwatch sw = Stopwatch.StartNew();
            workbook.Save("LargeWorkbook_Default.html", defaultOptions);
            sw.Stop();

            Console.WriteLine($"Export with HtmlCrossType.Default completed in {sw.Elapsed.TotalSeconds:F3} seconds.");

            // -----------------------------------------------------------------
            // Export using HtmlCrossType.Cross (set if property exists)
            // -----------------------------------------------------------------
            HtmlSaveOptions crossOptions = new HtmlSaveOptions(SaveFormat.Html);
            SetHtmlCrossTypeIfSupported(crossOptions, HtmlCrossType.Cross);

            sw.Restart();
            workbook.Save("LargeWorkbook_Cross.html", crossOptions);
            sw.Stop();

            Console.WriteLine($"Export with HtmlCrossType.Cross completed in {sw.Elapsed.TotalSeconds:F3} seconds.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Sets the HtmlCrossType property via reflection if the current Aspose.Cells version supports it.
    private static void SetHtmlCrossTypeIfSupported(HtmlSaveOptions options, HtmlCrossType value)
    {
        try
        {
            PropertyInfo prop = typeof(HtmlSaveOptions).GetProperty("HtmlCrossType", BindingFlags.Public | BindingFlags.Instance);
            if (prop != null && prop.CanWrite)
            {
                prop.SetValue(options, value);
            }
        }
        catch (Exception ex)
        {
            // If reflection fails, continue without setting the property.
            Console.WriteLine($"Warning: Unable to set HtmlCrossType ({value}). {ex.Message}");
        }
    }
}
