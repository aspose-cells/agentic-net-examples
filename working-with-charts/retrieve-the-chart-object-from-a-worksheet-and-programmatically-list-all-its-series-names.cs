// Title: Get Chart Series Names from a Worksheet using Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds sample data, builds a column chart, assigns series names from header cells, retrieves the chart via the worksheet's Charts collection, iterates the NSeries collection, and uses Series.DisplayName to output each series name before saving the file.
// Keywords: Aspose.Cells | C# chart series names | retrieve chart object | NSeries DisplayName | enumerate chart series | Excel chart API | get series name Aspose.Cells | Aspose.Cells chart example
// Common Searches: Aspose.Cells get chart series names C# | How to list series names from Excel chart using Aspose.Cells | Retrieve chart object from worksheet Aspose.Cells .NET | Enumerate NSeries and read DisplayName Aspose.Cells | C# code to read chart legend names Aspose.Cells
// Developer Intent: Obtain the chart object from a worksheet and extract the display names of all its data series.
// Use Cases: Validate that chart legends match source headers before publishing a workbook. | Generate a summary table of series names for reporting or analytics. | Synchronize chart series names with external systems or APIs. | Create dynamic UI elements (e.g., dropdowns) based on existing chart series. | Automate quality checks that ensure series names follow naming conventions.
// AI Prompts: Provide C# Aspose.Cells code that returns a List<string> of series DisplayName for a specified chart. | Show how to set series names from worksheet header cells and then read them back using Aspose.Cells. | Write a reusable method to fetch all series names from any chart in a workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartSeriesNames
{
    // This example creates a workbook, adds sample data, builds a column chart, assigns series names from header cells, retrieves the chart via the worksheet's Charts collection, iterates the NSeries collection, and uses Series.DisplayName to output each series name before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");

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

            // Set data source for the chart (two series)
            chart.NSeries.Add("B2:B4", true); // Series 1 values
            chart.NSeries.Add("C2:C4", true); // Series 2 values
            chart.NSeries.CategoryData = "A2:A4";

            // Optionally set series names from cells (so they have meaningful names)
            chart.NSeries[0].Name = "=Sheet1!$B$1";
            chart.NSeries[1].Name = "=Sheet1!$C$1";

            // Retrieve and list all series names
            Console.WriteLine("Chart Series Names:");
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                Series series = chart.NSeries[i];
                // Use the DisplayName property which reflects the name shown on the chart
                string seriesName = series.DisplayName;
                Console.WriteLine($"Series {i + 1}: {seriesName}");
            }

            // Save the workbook
            workbook.Save("ChartSeriesNamesOutput.xlsx");
        }
    }
}
