// Title: Add a Linear Trendline with Equation to a Line Series in a Combo Chart using Aspose.Cells for .NET
// Description: This C# example creates a workbook, populates quarterly data, builds a combo chart (column series on the primary axis and line series on the secondary axis), adds a blue linear trendline to the line series, shows its regression equation on the chart, and saves the file as an XLSX workbook.
// Keywords: Aspose.Cells combo chart trendline | C# add trendline to line series | display trendline equation Aspose.Cells | secondary axis line chart Aspose.Cells | linear regression chart .NET | Excel chart trendline programmatically
// Common Searches: how to add a trendline to a combo chart using Aspose.Cells | show regression equation on secondary axis line series Aspose.Cells C# | create column‑line combo chart with trendline .NET | Aspose.Cells add linear trendline to chart programmatically
// Developer Intent: Programmatically insert a linear trendline into the secondary‑axis line series of a combo chart and display its equation on the chart.
// Use Cases: Quarterly sales report: columns for revenue, line for profit margin with trendline equation for forecasting. | Financial dashboard: volume bars combined with price index line, trendline highlights market direction. | KPI performance sheet: target values as columns, actual performance as a line, equation aids quick trend analysis.
// AI Prompts: Generate C# code with Aspose.Cells to add a polynomial (order 2) trendline to the second series of an existing combo chart and show both the equation and R‑squared value. | Explain how to customize a trendline’s color, thickness, and dash style on a line series in a combo chart using Aspose.Cells for .NET. | Provide a step‑by‑step guide to retrieve the equation string of a trendline after it has been added to a chart with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, populates quarterly data, builds a combo chart (column series on the primary axis and line series on the secondary axis), adds a blue linear trendline to the line series, shows its regression equation on the chart, and saves the file as an XLSX workbook.
    class AddTrendlineToComboChart
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a combo chart (column + line)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("ColumnSeries");
                sheet.Cells["C1"].PutValue("LineSeries");

                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["A5"].PutValue("Q4");

                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);
                sheet.Cells["B5"].PutValue(200);

                sheet.Cells["C2"].PutValue(30);
                sheet.Cells["C3"].PutValue(45);
                sheet.Cells["C4"].PutValue(60);
                sheet.Cells["C5"].PutValue(80);

                // Add a combo chart (column as primary, line as secondary)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Add column series (first series)
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries[0].Name = "ColumnSeries";

                // Add line series (second series) and plot it on the secondary axis
                chart.NSeries.Add("C2:C5", true);
                chart.NSeries[1].Name = "LineSeries";
                chart.NSeries[1].PlotOnSecondAxis = true;
                chart.NSeries[1].Type = ChartType.Line; // set secondary chart type

                // Add a linear trendline to the line series and display its equation
                int trendlineIdx = chart.NSeries[1].TrendLines.Add(TrendlineType.Linear);
                Trendline trendline = chart.NSeries[1].TrendLines[trendlineIdx];
                trendline.DisplayEquation = true;
                trendline.DisplayRSquared = false;
                trendline.Color = Color.Blue;

                // Save the workbook
                string outputPath = "ComboChartWithTrendline.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {System.IO.Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
