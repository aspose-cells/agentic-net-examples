// Title: Add a moving average trendline with equation to a line chart in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an Excel file, locates the first line chart, adds a trendline of type MovingAverage (period 3) to its first series, enables equation display, and saves the workbook. | Demonstrate how to employ reflection in Aspose.Cells to retrieve the Series.Trendlines collection and invoke its Add method when the Trendlines property is not directly exposed. | Modify the sample so that the trendline period is passed as a method argument and the generated equation text is written into a designated worksheet cell.
// Common Searches: Aspose.Cells C# add moving average trendline to line chart | display trendline equation in Excel chart using Aspose.Cells | use reflection to add trendline in Aspose.Cells when Trendlines property missing | set moving average period for chart series with Aspose.Cells .NET | programmatically add trendline to existing chart in Aspose.Cells
// Tags: Aspose.Cells trendline addition C# | line chart trendline equation Aspose.Cells | reflection access series trendlines Aspose.Cells | configure trendline period Excel .NET | programmatic chart editing Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads (or creates) an Excel workbook, ensures a line chart is present, uses reflection to add a MovingAverage trendline with a period of 3 and equation display to the first series, records the result in cell A1, and saves the modified file.
class Program
{
    static void Main()
    {
        // Define input and output file paths
        string inputPath = "Input.xlsx";
        string outputPath = "Output.xlsx";

        try
        {
            Workbook workbook;

            // If the input file does not exist, create a sample workbook with a line chart
            if (!File.Exists(inputPath))
            {
                workbook = CreateSampleWorkbookWithChart();
                workbook.Save(inputPath);
            }
            else
            {
                // Load the existing workbook
                workbook = new Workbook(inputPath);
            }

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart on the sheet
            if (sheet.Charts.Count == 0)
                throw new InvalidOperationException("No charts found on the first worksheet.");

            // Assume the first chart is a line chart
            Chart chart = sheet.Charts[0];

            // Optional safety check: enforce line chart type
            if (chart.Type != ChartType.Line)
                chart.Type = ChartType.Line;

            // Ensure the chart has at least one series
            if (chart.NSeries.Count == 0)
                throw new InvalidOperationException("The chart does not contain any series.");

            // Add a moving average trendline to the first series (if supported)
            Series series = chart.NSeries[0];
            bool trendlineAdded = false;

            try
            {
                // Use reflection to access Trendlines collection (may not exist in older versions)
                PropertyInfo trendlinesProp = series.GetType().GetProperty("Trendlines");
                if (trendlinesProp != null)
                {
                    object trendlinesObj = trendlinesProp.GetValue(series, null);
                    MethodInfo addMethod = trendlinesObj.GetType().GetMethod("Add", new[] { typeof(TrendlineType) });
                    if (addMethod != null)
                    {
                        object trendlineObj = addMethod.Invoke(trendlinesObj, new object[] { TrendlineType.MovingAverage });

                        // Set trendline properties via reflection
                        PropertyInfo periodProp = trendlineObj.GetType().GetProperty("Period");
                        periodProp?.SetValue(trendlineObj, 3);

                        PropertyInfo displayEqProp = trendlineObj.GetType().GetProperty("DisplayEquation");
                        displayEqProp?.SetValue(trendlineObj, true);

                        trendlineAdded = true;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Trendline could not be added: {ex.Message}");
            }

            // Record result in cell A1
            sheet.Cells["A1"].PutValue(trendlineAdded
                ? "Moving Average Trendline added"
                : "Trendline not supported");

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Creates a simple workbook with sample data and a line chart
    private static Workbook CreateSampleWorkbookWithChart()
    {
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Sample data
        ws.Cells["A1"].PutValue("Category");
        ws.Cells["B1"].PutValue("Value");
        ws.Cells["A2"].PutValue("Jan");
        ws.Cells["A3"].PutValue("Feb");
        ws.Cells["A4"].PutValue("Mar");
        ws.Cells["B2"].PutValue(10);
        ws.Cells["B3"].PutValue(20);
        ws.Cells["B4"].PutValue(15);

        // Add a line chart
        int chartIndex = ws.Charts.Add(ChartType.Line, 5, 0, 20, 5);
        Chart chart = ws.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries[0].Name = "Sample Series";

        return wb;
    }
}
