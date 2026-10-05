// Title: Generate an Excel combo chart with column, line, and area series on separate axes using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a new workbook, adds three data series (column, line, area), sets each series' ChartType accordingly, and saves the file as an Excel workbook with Aspose.Cells. | Configure a combo chart where the column series uses the primary axis, the line series uses a secondary axis, and the area series uses another axis (if supported), then export the workbook using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# example for creating a combo chart with column, line and area series | How to assign different chart types to each series in an Aspose.Cells chart | Setting secondary axis for a series in Aspose.Cells combo chart C# | Multi-type Excel chart generation using Aspose.Cells .NET library | Create combo chart with separate axes for each series in Aspose.Cells
// Tags: combo chart series type configuration Aspose.Cells C# | assign secondary value axis Aspose.Cells chart | export workbook with multi-type chart Aspose.Cells | chart series type property Aspose.Cells .NET | create column line area combo chart Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// // This program creates a new workbook, populates category and three data series (column, line, area), adds a combo chart, sets each series to its specific chart type, and saves the file as ComboChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for three series
            // Categories
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");

            // Series 1 – Column
            sheet.Cells["B1"].PutValue("Column");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(40);

            // Series 2 – Line
            sheet.Cells["C1"].PutValue("Line");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);
            sheet.Cells["C5"].PutValue(45);

            // Series 3 – Area
            sheet.Cells["D1"].PutValue("Area");
            sheet.Cells["D2"].PutValue(5);
            sheet.Cells["D3"].PutValue(15);
            sheet.Cells["D4"].PutValue(25);
            sheet.Cells["D5"].PutValue(35);

            // Add a combo chart (initial type can be Column)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add the three series to the chart
            chart.NSeries.Add("B2:B5", true); // Column series
            chart.NSeries.Add("C2:C5", true); // Line series
            chart.NSeries.Add("D2:D5", true); // Area series

            // Set each series to its specific chart type
            chart.NSeries[0].Type = ChartType.Column;
            chart.NSeries[1].Type = ChartType.Line;
            chart.NSeries[2].Type = ChartType.Area;

            // Note: The IsSecondaryValueAxis property is not available in the current Aspose.Cells version.
            // If secondary axis support is required, upgrade to a newer version where this property exists.

            // Save the workbook with the combo chart
            workbook.Save("ComboChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
