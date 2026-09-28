// Title: Read trendline type and order from the first chart series in an XLSX workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an XLSX file with Aspose.Cells, verifies the first worksheet contains a chart, accesses the first series, and uses reflection to read the Trendline.Type and Trendline.Order properties while handling missing API gracefully. | Provide a sample that enumerates the Trendlines collection of a chart series via reflection, logs each trendline's type and order, and includes error handling for files without charts or trendlines.
// Common Searches: aspocells get trendline type from chart series c# | c# read trendline order using Aspose.Cells | how to use reflection to access chart trendline properties in Aspose.Cells .NET | retrieve confidence interval values of a chart trendline with Aspose.Cells | check if trendline feature is available in current Aspose.Cells version
// Tags: Aspose.Cells read chart trendline properties | C# extract trendline type from XLSX chart | reflection based access to Aspose.Cells trendlines | handle missing trendline API in Aspose.Cells | log trendline order and type in Excel workbook

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;
using System.Reflection;

// The example loads an XLSX workbook, ensures a worksheet, chart, and series exist, then uses reflection to obtain the Trendlines collection from the first series. It enumerates the first trendline and prints its Type and Order, with comprehensive checks for absent charts, series, or unsupported trendline APIs.
class Program
{
    static void Main()
    {
        try
        {
            string filePath = "input.xlsx";

            // Verify the input file exists
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook contains no worksheets.");
                return;
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the first worksheet.");
                return;
            }

            Chart chart = worksheet.Charts[0];

            // Ensure the chart has at least one series
            if (chart.NSeries.Count == 0)
            {
                Console.WriteLine("The chart contains no series.");
                return;
            }

            // Access the first series
            var series = chart.NSeries[0];

            // Use reflection to obtain the Trendlines collection (avoids compile‑time dependency)
            PropertyInfo trendlinesProp = series.GetType().GetProperty("Trendlines");
            if (trendlinesProp == null)
            {
                Console.WriteLine("Trendline feature is not available in this Aspose.Cells version.");
                return;
            }

            var trendlinesObj = trendlinesProp.GetValue(series);
            if (trendlinesObj == null)
            {
                Console.WriteLine("No trendlines collection found.");
                return;
            }

            // Trendlines implements IEnumerable; retrieve the first trendline if present
            var enumerator = (trendlinesObj as System.Collections.IEnumerable)?.GetEnumerator();
            if (enumerator == null || !enumerator.MoveNext())
            {
                Console.WriteLine("No trendlines found in the series.");
                return;
            }

            var trendline = enumerator.Current;
            if (trendline == null)
            {
                Console.WriteLine("Trendline object is null.");
                return;
            }

            // Retrieve Trendline properties via reflection
            PropertyInfo typeProp = trendline.GetType().GetProperty("Type");
            PropertyInfo orderProp = trendline.GetType().GetProperty("Order");

            var typeValue = typeProp?.GetValue(trendline);
            var orderValue = orderProp?.GetValue(trendline);

            Console.WriteLine($"Trendline Type: {typeValue}");
            Console.WriteLine($"Trendline Order: {orderValue}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
