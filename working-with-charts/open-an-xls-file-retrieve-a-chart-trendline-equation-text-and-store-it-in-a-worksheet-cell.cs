// Title: Read an XLS workbook, extract the first chart's trendline equation, and write it to cell A1 using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an .xls workbook with Aspose.Cells, locates the first chart's first series, uses reflection to read the trendline's formula, and stores the result in cell A1 before saving. | Show a .NET example that accesses a chart's Trendlines collection via Aspose.Cells, extracts the displayed equation text, writes it to a worksheet cell, and persists the modified file.
// Common Searches: aspocells c# retrieve chart trendline text from xls | write chart trendline equation into worksheet cell using Aspose.Cells | how to access Trendlines collection in Aspose.Cells chart series | extract first series trendline formula from Excel file with Aspose.Cells | save trendline equation to cell A1 in .NET
// Tags: Aspose.Cells read XLS workbook | Aspose.Cells chart trendline extraction | Aspose.Cells write cell value | Aspose.Cells reflection Trendlines access | C# chart series trendline retrieval

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample verifies the input .xls file, loads it with Aspose.Cells, accesses the first worksheet's first chart and its first series, uses reflection to obtain the Trendlines collection and the formula of the first trendline, writes the extracted equation (or a placeholder) into cell A1, and saves the updated workbook as output.xls while handling potential errors.
class TrendlineExtractor
{
    static void Main()
    {
        const string inputPath = "input.xls";
        const string outputPath = "output.xls";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            Chart chart = worksheet.Charts[0];

            // Ensure the chart has at least one series
            if (chart.NSeries.Count == 0)
            {
                Console.WriteLine("The chart does not contain any series.");
                return;
            }

            // Access the first series of the chart
            Series series = chart.NSeries[0];

            string equation = "Trendline not found";

            try
            {
                // Use reflection to access Trendlines collection (handles API variations)
                var trendlinesProp = series.GetType().GetProperty("Trendlines");
                if (trendlinesProp != null)
                {
                    var trendlines = trendlinesProp.GetValue(series);
                    var countProp = trendlines.GetType().GetProperty("Count");
                    int tlCount = (int)countProp.GetValue(trendlines);
                    if (tlCount > 0)
                    {
                        // Access first trendline (indexer property "Item")
                        var indexer = trendlines.GetType().GetProperty("Item");
                        var trendline = indexer.GetValue(trendlines, new object[] { 0 });

                        // Retrieve the formula/equation of the trendline
                        var formulaProp = trendline.GetType().GetProperty("Formula");
                        if (formulaProp != null)
                        {
                            equation = formulaProp.GetValue(trendline) as string ?? equation;
                        }
                    }
                }
            }
            catch (Exception reflEx)
            {
                Console.WriteLine($"Reflection error while extracting trendline: {reflEx.Message}");
            }

            // Store the equation (or placeholder) in cell A1
            worksheet.Cells["A1"].PutValue(equation);

            // Save the workbook with the new data
            workbook.Save(outputPath);
            Console.WriteLine($"Trendline equation saved to \"{outputPath}\" in cell A1.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
