// Title: Retrieve a Chart's Parent Worksheet Name with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, populates sample data, adds a column chart, and uses the Chart.Worksheet property to fetch the worksheet that hosts the chart. The worksheet name is written to the console for diagnostic purposes before the file is saved.
// Keywords: Aspose.Cells Chart.Worksheet | C# get chart parent worksheet | Aspose.Cells chart diagnostics | retrieve worksheet name from chart | .NET chart worksheet reference | Aspose.Cells example logging
// Common Searches: Aspose.Cells get worksheet of a chart | Chart.Worksheet property C# | how to log chart's parent sheet in Aspose.Cells | retrieve chart worksheet name Aspose.Cells .NET | diagnostic logging for charts Aspose.Cells
// Developer Intent: Identify the worksheet that contains a specific chart and output its name.
// Use Cases: Confirm that a newly created chart resides on the intended worksheet during development. | Create a diagnostic log that lists each chart together with its parent worksheet name. | Validate chart placement before saving the workbook by comparing Chart.Worksheet.Name to an expected value.
// AI Prompts: Generate C# code to iterate over all charts in a workbook and log each chart's parent worksheet name using Aspose.Cells. | Show how to write the chart's worksheet name to a file instead of the console in a .NET application. | Explain how Chart.Worksheet can be combined with workbook.Save to ensure charts are saved on the correct sheets.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartWorksheetDemo
{
    // This example creates a workbook, populates sample data, adds a column chart, and uses the Chart.Worksheet property to fetch the worksheet that hosts the chart. The worksheet name is written to the console for diagnostic purposes before the file is saved.
    public class Program
    {
        public static void Main()
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
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the parent worksheet of the chart and log its name
            Console.WriteLine("Chart's parent worksheet name: " + chart.Worksheet.Name);

            // Save the workbook (output file name can be adjusted as needed)
            workbook.Save("ChartWorksheetDemo_out.xlsx");
        }
    }
}
