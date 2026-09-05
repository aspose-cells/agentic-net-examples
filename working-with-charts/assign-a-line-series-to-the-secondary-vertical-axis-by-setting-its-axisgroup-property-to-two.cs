// Title: Assign a line chart series to the secondary vertical axis using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a line chart with two data series in Aspose.Cells and moves the second series to the secondary Y‑axis using the PlotOnSecondAxis property. | Show how to set a custom title for the secondary value axis of a line chart in Aspose.Cells with C#.
// Common Searches: Aspose.Cells C# line chart secondary Y axis example | how to use PlotOnSecondAxis in Aspose.Cells for .NET | set custom title for secondary axis in Aspose.Cells chart | move chart series to secondary vertical axis Aspose.Cells C# | create line chart with primary and secondary axes using Aspose.Cells
// Tags: Aspose.Cells PlotOnSecondAxis property | C# line chart secondary axis | Aspose.Cells secondary value axis title | Aspose.Cells chart series secondary axis | Aspose.Cells generate XLSX line chart

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSecondaryAxisDemo
{
    // Demonstrates creating a workbook, adding a line chart with two series, assigning the second series to the secondary vertical axis via PlotOnSecondAxis, customizing the secondary axis title, and saving the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Jan");
            worksheet.Cells["A3"].PutValue("Feb");
            worksheet.Cells["A4"].PutValue("Mar");

            worksheet.Cells["B1"].PutValue("Primary Series");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            worksheet.Cells["C1"].PutValue("Secondary Series");
            worksheet.Cells["C2"].PutValue(100);
            worksheet.Cells["C3"].PutValue(200);
            worksheet.Cells["C4"].PutValue(300);

            // Add a line chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Add two series: first will stay on primary axis, second on secondary axis
            chart.NSeries.Add("B2:B4", true); // primary series
            chart.NSeries.Add("C2:C4", true); // secondary series
            chart.NSeries.CategoryData = "A2:A4";

            // Assign the second series to the secondary vertical axis
            // In Aspose.Cells this is done via the PlotOnSecondAxis property
            chart.NSeries[1].PlotOnSecondAxis = true;

            // Optionally customize the secondary axis (e.g., title)
            Axis secondaryAxis = chart.SecondValueAxis;
            secondaryAxis.Title.Text = "Secondary Axis";

            // Save the workbook to a file
            workbook.Save("LineSeriesOnSecondaryAxis.xlsx");
        }
    }
}
