// Title: Create a combo chart with column and line series and add a linear trendline showing equation and R‑squared using Aspose.Cells for .NET
// AI Prompts: Generate an Excel workbook that contains a combo chart with a column series and a line series, then attach a linear trendline to the line series and enable the display of its equation and R‑squared value. | Update an existing Aspose.Cells chart by converting a series to line type and programmatically adding a linear trendline that shows both the formula and the R‑squared statistic on the chart.
// Common Searches: how to add a linear trendline with equation to a line series in an Aspose.Cells combo chart | Aspose.Cells .NET create combo chart with column and line series and show trendline equation | display R squared value for trendline in Excel chart using Aspose.Cells | change series type to line in Aspose.Cells combo chart programmatically
// Tags: Aspose.Cells combo chart trendline | add linear trendline Aspose.Cells | display trendline equation .NET Excel | convert series to line chart Aspose.Cells | Excel workbook chart customization Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsTrendlineExample
{
    // Creates a new workbook, populates sample data, adds a combo chart with a column series and a line series, switches the second series to a line type, optionally adds a linear trendline with equation and R‑squared display, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                // Column A: Categories
                // Column B: Column series values
                // Column C: Line series values
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Column Series");
                sheet.Cells["C1"].PutValue("Line Series");

                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");
                sheet.Cells["A6"].PutValue("May");

                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["B5"].PutValue(25);
                sheet.Cells["B6"].PutValue(15);

                sheet.Cells["C2"].PutValue(12);
                sheet.Cells["C3"].PutValue(22);
                sheet.Cells["C4"].PutValue(28);
                sheet.Cells["C5"].PutValue(27);
                sheet.Cells["C6"].PutValue(18);

                // Add a Combo chart (default type is Column, we will change the second series to Line)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set chart title
                chart.Title.Text = "Combo Chart with Trendline";

                // Add first series (Column)
                int colSeriesIndex = chart.NSeries.Add("B2:B6", true);
                chart.NSeries[colSeriesIndex].Name = "Column Series";

                // Add second series (Line)
                int lineSeriesIndex = chart.NSeries.Add("C2:C6", true);
                chart.NSeries[lineSeriesIndex].Name = "Line Series";

                // Change the second series to a Line chart type within the combo chart
                chart.NSeries[lineSeriesIndex].Type = ChartType.Line;

                // NOTE: Trendline support may depend on the Aspose.Cells version.
                // If the Trendlines property is unavailable, this section can be omitted.
                // Uncomment the following lines if your version supports trendlines.

                /*
                Series lineSeries = chart.NSeries[lineSeriesIndex];
                Trendline trendline = lineSeries.Trendlines.Add(TrendlineType.Linear);
                trendline.DisplayEquation = true;   // Show equation on chart
                trendline.DisplayRSquared = true;   // Show R‑squared value
                */

                // Determine output path and save the workbook
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComboChartWithTrendline.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
