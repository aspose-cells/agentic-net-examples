// Title: Hide second series in an Aspose.Cells line chart using C# (IsFiltered = true)
// Description: Demonstrates how to create a workbook, add sample data, generate a line chart, and hide the second data series by setting its IsFiltered property to true with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | line chart | hide chart series | IsFiltered | chart series visibility | Aspose.Cells chart | filter series | Excel chart manipulation
// Common Searches: Aspose.Cells hide series line chart C# | Set IsFiltered property Aspose.Cells chart | Hide second series in Excel chart using Aspose.Cells | Aspose.Cells chart series visibility .NET | How to filter a series in Aspose.Cells chart
// Developer Intent: Programmatically hide a specific data series in a line chart so it does not appear in the rendered output.
// Use Cases: Create dashboards where optional data series can be toggled off for cleaner visuals. | Generate financial or sales reports that include hidden series for calculations but not display. | Build template workbooks where certain chart series are pre‑filtered and can be revealed later.
// AI Prompts: Show C# code to hide a series in an Aspose.Cells line chart using the IsFiltered property. | How can I toggle visibility of chart series in an existing Aspose.Cells workbook with .NET? | Explain steps to filter out multiple series in an Aspose.Cells chart and later unfilter them.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add sample data, generate a line chart, and hide the second data series by setting its IsFiltered property to true with Aspose.Cells for .NET.
    public class HideSecondSeriesInLineChart
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for two series
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Jan");
            worksheet.Cells["A3"].PutValue("Feb");
            worksheet.Cells["A4"].PutValue("Mar");

            worksheet.Cells["B1"].PutValue("Series1");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            worksheet.Cells["C1"].PutValue("Series2");
            worksheet.Cells["C2"].PutValue(15);
            worksheet.Cells["C3"].PutValue(25);
            worksheet.Cells["C4"].PutValue(35);

            // Add a line chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Add the two series to the chart
            chart.NSeries.Add("B2:B4", true); // Series1
            chart.NSeries.Add("C2:C4", true); // Series2
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the second series by marking it as filtered (not displayed)
            chart.NSeries[1].IsFiltered = true;

            // Save the workbook
            string outputPath = "HideSecondSeriesLineChart.xlsx";
            workbook.Save(outputPath);
        }
    }
}
