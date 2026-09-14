// Title: Apply a logarithmic scale to the X‑axis of a scatter chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, inserts a scatter chart, and sets the X‑axis to logarithmic scaling with Aspose.Cells. | Show how to modify the logarithm base of the category (X) axis in an Aspose.Cells scatter chart. | Demonstrate adding multiple data series to a scatter chart while keeping the X‑axis in log scale using Aspose.Cells.
// Common Searches: Aspose.Cells C# scatter chart logarithmic X axis example | set category axis to log scale in Aspose.Cells .NET | change log base of X axis for Excel scatter chart using Aspose.Cells | how to enable log scaling on X axis of a scatter plot in C# with Aspose | create Excel scatter chart with log‑scaled X values using Aspose.Cells
// Tags: Aspose.Cells scatter chart log axis | C# set category axis logarithmic | Aspose.Cells chart axis log base | Excel scatter plot logarithmic X axis .NET | Aspose.Cells workbook chart creation

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, populates X and Y data, adds a scatter chart, assigns the data ranges, enables logarithmic scaling on the category (X) axis (default base 10), optionally sets a custom log base, and saves the file as ScatterChart_LogScale.xlsx.
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

                // Populate sample data for the scatter chart (X values in column A, Y values in column B)
                sheet.Cells["A1"].PutValue("X");
                sheet.Cells["B1"].PutValue("Y");
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue(5);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue(10);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["A5"].PutValue(50);
                sheet.Cells["B5"].PutValue(40);
                sheet.Cells["A6"].PutValue(100);
                sheet.Cells["B6"].PutValue(50);

                // Add a scatter chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the chart title (optional)
                chart.Title.Text = "Scatter Chart with Logarithmic X‑Axis";

                // Add a series to the chart using the Y‑values range
                int seriesIndex = chart.NSeries.Add("B2:B6", true);
                // Retrieve the series (type is Series, not NSeries)
                Series series = chart.NSeries[seriesIndex];
                // Assign X‑values range
                series.XValues = "A2:A6";

                // Apply logarithmic scale to the X‑axis (Category Axis)
                chart.CategoryAxis.IsLogarithmic = true;
                // Optionally, set the base of the logarithm (default is 10)
                chart.CategoryAxis.LogBase = 10;

                // Save the workbook to a file
                string outputPath = "ScatterChart_LogScale.xlsx";
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
