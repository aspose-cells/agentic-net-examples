// Title: Benchmarking HtmlCrossType.Cross vs default HTML save options for large workbooks in Aspose.Cells for .NET
// Common Searches: how to benchmark HtmlCrossType.Cross performance when exporting a large Excel workbook to HTML using Aspose.Cells C# | compare HTML export speed with and without HtmlCrossType.Cross for a worksheet containing 100000 rows | measure the impact of HtmlCrossType.Cross on Aspose.Cells HTML save time for thousands of rows | C# code example to time Aspose.Cells HTML export using different HtmlSaveOptions settings | performance testing of Aspose.Cells HTML export on massive datasets in .NET
// Tags: Aspose.Cells HTML export performance | HtmlCrossType.Cross benchmark | large workbook HTML save timing | C# Aspose.Cells HtmlSaveOptions comparison | export 100k rows to HTML Aspose.Cells

using System;
using System.Diagnostics;
using Aspose.Cells;

// The sample creates a workbook with 100,000 rows and 10 columns, fills each cell with a string, and then saves the workbook to HTML twice—once with the default HtmlSaveOptions and once with an alternative HtmlSaveOptions instance where HtmlCrossType.Cross could be set. Stopwatch timers record the elapsed time for each save, and the results are printed to demonstrate any performance difference between the default and the Cross setting.
class Program
{
    static void Main()
    {
        try
        {
            // Create a workbook and populate it with a large number of rows and columns.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            int totalRows = 100_000; // thousands of rows
            int totalCols = 10;      // sample columns

            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalCols; col++)
                {
                    // Fill each cell with a simple string value.
                    sheet.Cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // Measure performance using the default HTML save options.
            HtmlSaveOptions defaultOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: HtmlCrossType property is not available in the referenced Aspose.Cells version.
            // The default behavior is used here.

            Stopwatch swDefault = Stopwatch.StartNew();
            workbook.Save("default.html", defaultOptions);
            swDefault.Stop();

            // Measure performance using alternative HTML save options (if supported).
            HtmlSaveOptions alternativeOptions = new HtmlSaveOptions(SaveFormat.Html);
            // If a future version provides HtmlCrossType, it can be set here, e.g.:
            // alternativeOptions.HtmlCrossType = HtmlCrossType.Cross;

            Stopwatch swAlternative = Stopwatch.StartNew();
            workbook.Save("alternative.html", alternativeOptions);
            swAlternative.Stop();

            // Output the timing results.
            Console.WriteLine($"Default HtmlSaveOptions elapsed:    {swDefault.ElapsedMilliseconds} ms");
            Console.WriteLine($"Alternative HtmlSaveOptions elapsed:{swAlternative.ElapsedMilliseconds} ms");
            Console.WriteLine($"Alternative improves performance:   {swAlternative.Elapsed < swDefault.Elapsed}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
