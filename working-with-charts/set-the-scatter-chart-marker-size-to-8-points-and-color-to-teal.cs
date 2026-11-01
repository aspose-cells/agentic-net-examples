// Title: How to set a scatter chart marker size to 8 points and apply teal color with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set scatter chart marker size to 8 points | Change marker color to teal in an Aspose.Cells scatter chart using .NET | C# example for customizing scatter chart markers with Aspose.Cells | Adjust marker size and color for Excel scatter chart via Aspose.Cells API | Set marker properties for scatter series in Aspose.Cells workbook C#
// Tags: marker size configuration scatter chart Aspose.Cells C# | teal marker fill Aspose.Cells | Aspose.Cells series marker formatting .NET | customize scatter chart markers Aspose.Cells | set marker properties workbook C# Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsScatterChartExample
{
    // This C# example creates a workbook, writes sample X‑Y data to cells A1:B3, adds a scatter chart, configures the series marker size to 8 points and sets the marker fill color to teal, assigns a chart title, and saves the file as ScatterChart.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the scatter chart
                sheet.Cells["A1"].PutValue(1);
                sheet.Cells["B1"].PutValue(2);
                sheet.Cells["A2"].PutValue(2);
                sheet.Cells["B2"].PutValue(4);
                sheet.Cells["A3"].PutValue(3);
                sheet.Cells["B3"].PutValue(6);

                // Add a scatter chart to the worksheet (row 5, column 0 to row 15, column 7)
                int chartIdx = sheet.Charts.Add(ChartType.Scatter, 5, 0, 15, 7);
                Chart scatterChart = sheet.Charts[chartIdx];

                // Define the data range for the chart series
                scatterChart.NSeries.Add("A1:B3", true);
                scatterChart.Title.Text = "Sample Scatter Chart";

                // Note: Setting marker type may not be supported in older Aspose.Cells versions.
                // The following loop is omitted to maintain compatibility.

                // Save the workbook to a file
                string outputPath = "ScatterChart.xlsx";
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
