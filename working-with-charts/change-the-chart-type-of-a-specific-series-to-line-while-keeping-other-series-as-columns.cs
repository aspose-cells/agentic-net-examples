// Title: Aspose.Cells for .NET: Change a single chart series to Line while keeping other series as Columns (C#)
// Description: C# example that creates a workbook, adds three data series, builds a column chart, then converts only the third series to a line chart by setting NSeries[2].Type = ChartType.Line. The other series remain columns and the workbook is saved as SeriesTypeChanged.xlsx.
// Keywords: Aspose.Cells | .NET chart example | C# Aspose.Cells series type | ChartType.Line | combo column line chart | NSeries Type property | programmatic chart modification | sample code GitHub | Aspose.Cells tutorial
// Common Searches: Aspose.Cells change one series to line chart | C# combo chart column and line Aspose.Cells | set individual series type Aspose.Cells | ChartType.Line example Aspose.Cells .NET | how to keep other series as columns Aspose.Cells
// Developer Intent: Convert a specific series in an existing column chart to a line series while leaving the remaining series unchanged.
// Use Cases: Display monthly sales as columns with a target trend line on the same chart. | Show production volume as bars and efficiency percentage as a line for performance dashboards. | Combine revenue columns with a moving‑average line in financial reporting.
// AI Prompts: Generate C# code that changes the second series of a column chart to a spline series using Aspose.Cells. | Explain how to let users select a series at runtime and switch its chart type in an Aspose.Cells workbook. | Show how to iterate through a chart's NSeries collection and assign different ChartType values to each series.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSeriesTypeDemo
{
    // C# example that creates a workbook, adds three data series, builds a column chart, then converts only the third series to a line chart by setting NSeries[2].Type = ChartType.Line. The other series remain columns and the workbook is saved as SeriesTypeChanged.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for three series
            // Category column
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            // Series 1 (will stay as Column)
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Series 2 (will stay as Column)
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Series 3 (will be changed to Line)
            sheet.Cells["D1"].PutValue("Series3");
            sheet.Cells["D2"].PutValue(12);
            sheet.Cells["D3"].PutValue(22);
            sheet.Cells["D4"].PutValue(32);

            // Add a column chart that initially contains all series as columns
            int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data range for all four series
            chart.NSeries.Add("B2:D4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Change the type of the third series (index 2) to Line
            // Other series (index 0 and 1) remain as Column
            chart.NSeries[2].Type = ChartType.Line;

            // Optional: give the series a distinct name to verify the change
            chart.NSeries[2].Name = "Series3 (Line)";

            // Save the workbook
            workbook.Save("SeriesTypeChanged.xlsx");
        }
    }
}
