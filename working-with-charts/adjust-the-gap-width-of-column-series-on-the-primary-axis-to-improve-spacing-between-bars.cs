// Title: Adjust Column Series GapWidth in Aspose.Cells .NET Chart (C#) to Control Bar Spacing
// Description: Creates a workbook, adds sample data, inserts a 2‑D column chart, binds values and categories, sets the first series' GapWidth to 150 % (range 0‑500) on the primary axis, and saves the file as AdjustedGapWidth.xlsx.
// Keywords: Aspose.Cells | C# | .NET | column chart | GapWidth property | bar spacing | primary axis | Excel chart formatting | adjust column series gap | set series gap width
// Common Searches: Aspose.Cells change column series gap width C# | how to increase spacing between bars in Aspose.Cells chart | set GapWidth for column chart series .NET | adjust column chart bar width programmatically Aspose | gap width range Aspose.Cells column chart
// Developer Intent: Programmatically set the GapWidth of a column series to modify bar spacing in an Aspose.Cells chart.
// Use Cases: Design column charts with custom bar spacing for clearer visual comparison. | Generate Excel reports where dense data requires wider gaps between columns. | Apply corporate style guidelines by programmatically adjusting column gaps.
// AI Prompts: Provide C# code that sets the GapWidth of a column series in an Aspose.Cells chart. | Show how to apply different GapWidth values to multiple series in a stacked column chart using Aspose.Cells. | Explain the visual impact of GapWidth percentages on column charts in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, inserts a 2‑D column chart, binds values and categories, sets the first series' GapWidth to 150 % (range 0‑500) on the primary axis, and saves the file as AdjustedGapWidth.xlsx.
    public class AdjustColumnSeriesGapWidth
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the column chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a 2‑D column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories

                // Adjust the gap width of the first (and only) column series on the primary axis
                // GapWidth is a percentage of the column width (0‑500). Here we set it to 150%
                chart.NSeries[0].GapWidth = 150;

                // Save the workbook
                workbook.Save("AdjustedGapWidth.xlsx");
                Console.WriteLine("Workbook saved as AdjustedGapWidth.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            AdjustColumnSeriesGapWidth.Run();
        }
    }
}
