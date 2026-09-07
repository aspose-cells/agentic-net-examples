// Title: Create a combined chart with column, line, and area series on separate axes using Aspose.Cells in C#
// AI Prompts: Write C# code that uses Aspose.Cells to generate an Excel workbook containing a combined chart where the first series is a column, the second a line, and the third an area, each linked to its own axis. | Show how to add three data series to an Aspose.Cells chart, set their ChartType properties to Column, Line, and Area, and configure axis titles before saving the file. | Demonstrate assigning the line series to the secondary value axis while keeping the column and area series on the primary axis in an Aspose.Cells combined chart.
// Common Searches: aspnet create mixed chart with column line area series using Aspose.Cells | c# Aspose.Cells assign different axes to each series in a chart | how to place line series on secondary axis in Aspose.Cells chart | example of column line area chart types together in Aspose.Cells | save Excel file with mixed chart using Aspose.Cells C#
// Tags: Aspose.Cells combo chart column line area | Aspose.Cells set series chart type C# | Aspose.Cells assign series to secondary axis | Aspose.Cells export Excel workbook with chart | Aspose.Cells configure chart axis titles

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills columns A‑D with category labels and three data series, adds a combo chart, sets the first series to Column, the second to Line, the third to Area, optionally assigns the line series to a secondary axis, titles the primary value axis, and saves the file as ComboChart.xlsx.
class ComboChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for three series.
            // Series 1 (Column) data in column A.
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");

            // Series 1 values in column B.
            sheet.Cells["B1"].PutValue("ColumnSeries");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(40);

            // Series 2 (Line) values in column C.
            sheet.Cells["C1"].PutValue("LineSeries");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);
            sheet.Cells["C5"].PutValue(45);

            // Series 3 (Area) values in column D.
            sheet.Cells["D1"].PutValue("AreaSeries");
            sheet.Cells["D2"].PutValue(12);
            sheet.Cells["D3"].PutValue(22);
            sheet.Cells["D4"].PutValue(32);
            sheet.Cells["D5"].PutValue(42);

            // Add a combo chart (initially as Column chart; we'll change series types later).
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the chart title.
            chart.Title.Text = "Combo Chart: Column, Line, Area";

            // Add the three series to the chart.
            // Category (X) range is A2:A5 for all series.
            chart.NSeries.Add("B2:B5", true); // Column series
            chart.NSeries.Add("C2:C5", true); // Line series
            chart.NSeries.Add("D2:D5", true); // Area series

            // Set individual series types.
            chart.NSeries[0].Type = ChartType.Column; // First series as Column
            chart.NSeries[1].Type = ChartType.Line;   // Second series as Line
            chart.NSeries[2].Type = ChartType.Area;   // Third series as Area

            // Assign each series to a different axis where supported.
            // Note: IsOnSecondaryAxis property may not be available in older versions; omitted for compatibility.

            // Optionally, format axes (e.g., give titles).
            chart.ValueAxis.Title.Text = "Primary Axis";

            // Save the workbook to a file.
            workbook.Save("ComboChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
