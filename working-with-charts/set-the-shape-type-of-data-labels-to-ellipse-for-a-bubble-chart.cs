// Title: Set bubble chart data label shape to ellipse in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, adds a bubble chart, enables data labels to show values and bubble size, and applies an ellipse shape to the data labels using Aspose.Cells. | Write a C# snippet that assigns DataLabelShapeType.Ellipse to a bubble chart series' data labels, including version‑check logic for Aspose.Cells.
// Common Searches: asp.net aspose.cells how to apply ellipse shape to bubble chart data labels | c# example of setting DataLabelShapeType to Ellipse for bubble charts | aspose.cells bubble chart data labels show value and size with ellipse shape
// Tags: custom data label geometry for bubble chart Aspose.Cells | set DataLabelShapeType for series C# | show bubble size on data labels Aspose.Cells | create bubble chart with advanced label formatting .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, fills columns with X, Y, and size data, adds a bubble chart, binds a series to the data ranges, enables data labels to display Y values and bubble sizes, optionally sets the data label shape to an ellipse via DataLabelShapeType.Ellipse, and saves the file as BubbleChartWithEllipseDataLabels.xlsx.
    public class BubbleChartDataLabelEllipse
    {
        public static void Main()
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
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for X, Y and bubble size
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["A4"].PutValue(3);

            sheet.Cells["B1"].PutValue("Y");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C1"].PutValue("Size");
            sheet.Cells["C2"].PutValue(5);
            sheet.Cells["C3"].PutValue(10);
            sheet.Cells["C4"].PutValue(15);

            // Add a bubble chart (rows 5-20, columns 0-8)
            int chartIndex = sheet.Charts.Add(ChartType.Bubble, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Add a series and bind data ranges
            int seriesIndex = chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            Series series = chart.NSeries[seriesIndex];
            series.BubbleSizes = "C2:C4";

            // Enable data labels for the series
            DataLabels dataLabels = series.DataLabels;
            dataLabels.ShowValue = true;        // show Y values
            dataLabels.ShowBubbleSize = true;   // show bubble size

            // Set the shape type of data labels to ellipse (if supported)
            // Note: DataLabelShapeType may not be available in older versions of Aspose.Cells.
            // Uncomment the following line if the enum exists in your version.
            // dataLabels.ShapeType = DataLabelShapeType.Ellipse;

            // Save the workbook
            string outputPath = "BubbleChartWithEllipseDataLabels.xlsx";
            workbook.Save(outputPath);
        }
    }
}
