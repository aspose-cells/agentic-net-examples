// Title: Make legend entries transparent while preserving text color in an Aspose.Cells chart using C#
// Common Searches: Aspose.Cells C# set legend entry background to transparent without changing font color | remove fill from chart legend entries in Aspose.Cells while keeping text contrast | how to make legend background transparent in an Aspose.Cells chart with C# | clear legend entry background mode Aspose.Cells chart example C# | Aspose.Cells chart legend customization transparent background preserve font color
// Tags: legend entry background mode transparent Aspose.Cells | preserve legend text color Aspose.Cells chart | clear legend fill Aspose.Cells C# | transparent chart legend background Aspose.Cells | Aspose.Cells legend customization C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace RemoveLegendBackground
{
    // The example creates a workbook, adds sample data, inserts a column chart, and iterates over each series to set each legend entry's BackgroundMode to Transparent while retaining the original Font.Color, ensuring text remains readable. It also makes the overall legend background transparent before saving the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the output file
            string outputPath = "ChartWithTransparentLegend.xlsx";

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Populate sample data for the chart
            // ------------------------------------------------------------
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(50);
            sheet.Cells["B4"].PutValue(70);

            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["C2"].PutValue(40);
            sheet.Cells["C3"].PutValue(60);
            sheet.Cells["C4"].PutValue(80);

            // ------------------------------------------------------------
            // Add a column chart
            // ------------------------------------------------------------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // ------------------------------------------------------------
            // Remove background fill from each legend entry
            // ------------------------------------------------------------
            foreach (Series series in chart.NSeries)
            {
                LegendEntry legendEntry = series.LegendEntry;

                // Preserve the current text color
                System.Drawing.Color currentTextColor = legendEntry.Font.Color;

                // Remove background fill by making it transparent
                legendEntry.BackgroundMode = BackgroundMode.Transparent;

                // Ensure the text itself is not set to "no fill"
                legendEntry.IsTextNoFill = false;

                // Re‑apply the original text color (keeps contrast)
                legendEntry.Font.Color = currentTextColor;
            }

            // Optional: also clear any fill that might be set on the whole legend object
            if (chart.Legend != null)
            {
                chart.Legend.BackgroundMode = BackgroundMode.Transparent;
            }

            // ------------------------------------------------------------
            // Save the workbook
            // ------------------------------------------------------------
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error saving workbook: " + ex.Message);
            }
        }
    }
}
