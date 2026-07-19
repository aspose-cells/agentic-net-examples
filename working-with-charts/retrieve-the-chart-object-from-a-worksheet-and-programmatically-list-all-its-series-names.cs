// Title: Aspose.Cells for .NET – Retrieve a Chart and List All Series Names
// Description: Creates a workbook, adds sample data, builds a column chart, assigns series names from header cells, then iterates the chart's NSeries collection to output each series name (or fallback display name) and saves the file.
// Keywords: Aspose.Cells chart series names | list chart series .NET | retrieve chart object Aspose.Cells | enumerate NSeries Aspose.Cells | C# Aspose.Cells chart legend | extract series titles Excel workbook
// Common Searches: how to get series names from an Aspose.Cells chart | Aspose.Cells enumerate chart series in C# | read chart legend titles using Aspose.Cells | C# list NSeries names Aspose.Cells | extract chart series titles programmatically
// Developer Intent: Obtain the chart object from a worksheet and programmatically output the names of every series it contains.
// Use Cases: Verify that series titles are correctly linked to header cells before publishing a report. | Generate a custom legend or UI dropdown populated with chart series names at runtime. | Log or assert series names in automated tests to ensure data integrity.
// AI Prompts: Write C# code with Aspose.Cells that returns a List<string> of all series names from every chart on a worksheet. | Show how to check a series Name property for a cell reference and use DisplayName as a fallback when the Name is empty. | Explain how to extend the example to handle multiple charts on the same sheet and collect their series names into a dictionary.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartSeriesNames
{
    // Creates a workbook, adds sample data, builds a column chart, assigns series names from header cells, then iterates the chart's NSeries collection to output each series name (or fallback display name) and saves the file.
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

            worksheet.Cells["B1"].PutValue("Series 1");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            worksheet.Cells["C1"].PutValue("Series 2");
            worksheet.Cells["C2"].PutValue(15);
            worksheet.Cells["C3"].PutValue(25);
            worksheet.Cells["C4"].PutValue(35);

            // Add a chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart series
            chart.NSeries.Add("B2:C4", true);          // Y values for both series
            chart.NSeries.CategoryData = "A2:A4";      // X axis categories

            // Optionally assign names to the series from cells
            chart.NSeries[0].Name = "=B1";
            chart.NSeries[1].Name = "=C1";

            // Retrieve the chart object (already have it) and list all series names
            Console.WriteLine("Series names in the chart:");
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                Series series = chart.NSeries[i];
                // Use the Name property (if set) or fallback to DisplayName
                string seriesName = !string.IsNullOrEmpty(series.Name) ? series.Name : series.DisplayName;
                Console.WriteLine($"Series {i + 1}: {seriesName}");
            }

            // Save the workbook
            workbook.Save("ChartSeriesNames.xlsx");
        }
    }
}
