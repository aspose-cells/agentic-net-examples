// Title: Benchmark HTML export speed of a large workbook using HtmlCrossType.Cross vs HtmlCrossType.Default in Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a 20,000‑row worksheet, saves it to HTML with HtmlSaveOptions set to the Cross mode, and logs the elapsed milliseconds. | Write a C# snippet that re‑exports the same workbook to HTML using the Default mode and prints both export durations for side‑by‑side comparison. | Provide a C# example that measures and displays the performance difference between Cross and Default HTML export modes when converting a large workbook with Aspose.Cells.
// Common Searches: Aspose.Cells how to benchmark HTML export time for a large workbook | C# compare HtmlCrossType.Cross and HtmlCrossType.Default performance | measure Aspose.Cells HTML save speed for 20000 rows | timing HTML conversion with Aspose.Cells SaveOptions | performance test Aspose.Cells HTML export large worksheet
// Tags: Aspose.Cells HTML export performance testing | Cross vs Default HTML save mode comparison | large workbook HTML conversion timing | C# benchmark Aspose.Cells HtmlSaveOptions | measure HTML export latency Aspose.Cells

using System;
using System.Diagnostics;
using Aspose.Cells;

// // Creates a 20,000‑row by 50‑column workbook, fills it with sample data, then measures and prints the elapsed time for saving the workbook to HTML using the Cross mode and the Default mode via HtmlSaveOptions.
class HtmlExportTiming
{
    static void Main()
    {
        try
        {
            // Create a large workbook with sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            int totalRows = 20000;   // adjust size for a large workbook
            int totalCols = 50;

            // Populate cells with dummy data
            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalCols; col++)
                {
                    cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // Measure export time using default HtmlCrossType (Cross is not available in older versions)
            HtmlSaveOptions optionsCross = new HtmlSaveOptions(SaveFormat.Html);
            Stopwatch swCross = Stopwatch.StartNew();
            workbook.Save("LargeWorkbook_Cross.html", optionsCross);
            swCross.Stop();

            // Measure export time using default HtmlCrossType (Default)
            HtmlSaveOptions optionsDefault = new HtmlSaveOptions(SaveFormat.Html);
            Stopwatch swDefault = Stopwatch.StartNew();
            workbook.Save("LargeWorkbook_Default.html", optionsDefault);
            swDefault.Stop();

            // Output the measured times
            Console.WriteLine($"Export time with default (Cross) : {swCross.ElapsedMilliseconds} ms");
            Console.WriteLine($"Export time with default (Default) : {swDefault.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
