// Title: Hide the horizontal (category) axis gridlines in a C# Aspose.Cells scatter chart
// AI Prompts: Generate C# code with Aspose.Cells that creates a scatter chart and sets CategoryAxis.MajorGridLines.IsVisible to false. | Show how to programmatically toggle the visibility of chart axis gridlines based on a runtime condition in a .NET application using Aspose.Cells. | Provide an example that hides both horizontal and vertical major gridlines in an Aspose.Cells chart and saves the workbook.
// Common Searches: Aspose.Cells C# hide category axis gridlines in scatter plot | remove major gridlines from Excel chart using Aspose.Cells .NET | how to turn off horizontal axis lines in Aspose.Cells chart programmatically | C# Aspose.Cells scatter chart without gridlines example | disable chart axis gridlines in generated Excel file with Aspose.Cells
// Tags: scatter chart hide major gridlines Aspose.Cells | Aspose.Cells chart major gridlines visibility C# | disable horizontal axis gridlines Excel chart Aspose | Aspose.Cells axis formatting hide gridlines | C# generate scatter plot without gridlines Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ScatterPlotExample
{
    // The sample creates a new workbook, fills columns A and B with X/Y data, adds a scatter chart, sets the CategoryAxis major gridlines to invisible for a cleaner look, and saves the file as ScatterPlot_HideHorizontalGridlines.xlsx.
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

                // Populate sample data for the scatter plot (X values in column A, Y values in column B)
                sheet.Cells["A1"].PutValue("X");
                sheet.Cells["B1"].PutValue("Y");
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue(2);
                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue(4);
                sheet.Cells["A4"].PutValue(3);
                sheet.Cells["B4"].PutValue(6);
                sheet.Cells["A5"].PutValue(4);
                sheet.Cells["B5"].PutValue(8);
                sheet.Cells["A6"].PutValue(5);
                sheet.Cells["B6"].PutValue(10);

                // Add a scatter chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Scatter, 7, 0, 27, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart series
                chart.NSeries.Add("B2:B6", true);
                chart.NSeries[0].XValues = "A2:A6";

                // Hide the horizontal axis (Category axis) major gridlines for a cleaner appearance
                chart.CategoryAxis.MajorGridLines.IsVisible = false;

                // Optionally, also hide the vertical axis gridlines if desired
                // chart.ValueAxis.MajorGridLines.IsVisible = false;

                // Save the workbook to a file
                string outputPath = "ScatterPlot_HideHorizontalGridlines.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
