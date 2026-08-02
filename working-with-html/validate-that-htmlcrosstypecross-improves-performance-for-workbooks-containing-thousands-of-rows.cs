// Title: Benchmark HtmlCrossType.Cross vs Default for Large HTML Export in Aspose.Cells .NET
// Description: This C# example builds a 10,000‑row, 5‑column workbook filled with overflow strings, saves it twice—once with HtmlSaveOptions.HtmlCrossStringType set to Default and once to Cross—and reports the elapsed milliseconds. Use it to verify whether the Cross option speeds up HTML conversion for massive worksheets.
// Keywords: Aspose.Cells HtmlCrossType performance | HtmlCrossType.Cross speed | HTML export large workbook .NET | Aspose.Cells benchmark HTML save | Cross vs Default HtmlCrossStringType
// Common Searches: Aspose.Cells HtmlCrossType.Cross performance test | HTML export speed comparison Default vs Cross | benchmark large worksheet HTML save Aspose | does HtmlCrossType.Cross improve export time | measure Aspose.Cells HTML conversion latency
// Developer Intent: Determine if setting HtmlCrossStringType to Cross reduces the time required to export a workbook with thousands of rows to HTML.
// Use Cases: Run a quick performance test before selecting the HtmlCrossStringType for a production reporting service. | Validate the benefit of the Cross option when generating HTML reports from datasets that cause cell overflow. | Provide empirical data for capacity planning of a web API that repeatedly converts large spreadsheets to HTML.
// AI Prompts: Generate C# code that repeats the HTML save with HtmlCrossType.Cross 10 times and calculates the average duration. | Explain how HtmlCrossType.Cross handles overflow text and why it can be faster than the Default mode. | Create a snippet that logs both Default and Cross export times to a CSV file for later analysis.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsPerformanceDemo
{
    // This C# example builds a 10,000‑row, 5‑column workbook filled with overflow strings, saves it twice—once with HtmlSaveOptions.HtmlCrossStringType set to Default and once to Cross—and reports the elapsed milliseconds. Use it to verify whether the Cross option speeds up HTML conversion for massive worksheets.
    public class HtmlCrossTypePerformance
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate the worksheet with a large amount of data (e.g., 10,000 rows, 5 columns)
                int totalRows = 10000;
                int totalColumns = 5;
                for (int row = 0; row < totalRows; row++)
                {
                    for (int col = 0; col < totalColumns; col++)
                    {
                        // Use a string that will overflow the cell width to trigger cross‑cell behavior
                        cells[row, col].PutValue($"Row{row}_Col{col}_LongTextThatWillOverflow");
                    }
                }

                // Prepare HTML save options with the default cross type (Default)
                HtmlSaveOptions defaultOptions = new HtmlSaveOptions
                {
                    HtmlCrossStringType = HtmlCrossType.Default
                };

                // Measure time for saving with Default cross type
                Stopwatch swDefault = Stopwatch.StartNew();
                workbook.Save("LargeWorkbook_Default.html", defaultOptions);
                swDefault.Stop();

                // Prepare HTML save options with the performance‑optimized Cross type
                HtmlSaveOptions crossOptions = new HtmlSaveOptions
                {
                    HtmlCrossStringType = HtmlCrossType.Cross
                };

                // Measure time for saving with Cross cross type
                Stopwatch swCross = Stopwatch.StartNew();
                workbook.Save("LargeWorkbook_Cross.html", crossOptions);
                swCross.Stop();

                // Output the measured times
                Console.WriteLine($"Saving with HtmlCrossType.Default took: {swDefault.ElapsedMilliseconds} ms");
                Console.WriteLine($"Saving with HtmlCrossType.Cross   took: {swCross.ElapsedMilliseconds} ms");
                Console.WriteLine("Performance improvement can be observed if the Cross option is faster.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            HtmlCrossTypePerformance.Run();
        }
    }
}
