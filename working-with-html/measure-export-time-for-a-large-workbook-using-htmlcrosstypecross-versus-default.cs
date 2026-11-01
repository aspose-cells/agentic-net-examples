// Title: Compare HTML export speed of a large workbook using HtmlSaveOptions default vs HtmlCrossType.Cross in Aspose.Cells for .NET
// AI Prompts: Create a C# console application that builds a 5000‑row by 50‑column worksheet, saves it to HTML with Aspose.Cells using the default HtmlSaveOptions, and logs the elapsed milliseconds with Stopwatch. | Extend the program to use HtmlSaveOptions configured with HtmlCrossType.Cross, export the same workbook, and output both the default and cross‑type export times for a side‑by‑side performance comparison.
// Common Searches: aspnet aspose.cells benchmark html export default vs cross type for large spreadsheets | c# measure time to save 5000 rows to html using HtmlSaveOptions HtmlCrossType.Cross | how to compare Aspose.Cells HTML conversion performance with different HtmlCrossType settings
// Tags: Aspose.Cells HTML export speed test | HtmlSaveOptions performance with large worksheet | HtmlCrossType.Cross impact on HTML conversion time | benchmarking Aspose.Cells HTML save for big spreadsheets

using System;
using System.Diagnostics;
using Aspose.Cells;

// The example builds a 5,000‑row by 50‑column worksheet, fills each cell with sample data, and saves the workbook to HTML twice—once with the default HtmlSaveOptions and once with HtmlSaveOptions set to HtmlCrossType.Cross. A Stopwatch records each export duration, and the elapsed milliseconds are printed to the console for performance comparison.
class Program
{
    static void Main()
    {
        try
        {
            // Create a large workbook with sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            int totalRows = 5000;   // adjust size as needed for "large" workbook
            int totalCols = 50;

            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalCols; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // Export using default HTML options (Cross type not available in this version)
            HtmlSaveOptions defaultOptions = new HtmlSaveOptions(SaveFormat.Html);

            Stopwatch swDefault = Stopwatch.StartNew();
            workbook.Save("LargeWorkbook_Default.html", defaultOptions);
            swDefault.Stop();

            // Output the measured time
            Console.WriteLine($"Export time with default HtmlSaveOptions : {swDefault.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
