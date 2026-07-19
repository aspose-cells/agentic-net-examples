// Title: Load an XLSX workbook and retrieve the first chart using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to open an existing XLSX file with the Workbook(string) constructor, access the first worksheet, obtain its ChartCollection, verify that a chart exists, and fetch the chart at index 0. The sample prints the chart's Name, Type, and the parent worksheet name.
// Keywords: Aspose.Cells load workbook C# | Workbook(string) constructor | first worksheet chart | ChartCollection Aspose.Cells | retrieve chart index 0 | read chart name and type | .NET chart API | C# Excel chart extraction
// Common Searches: How to load an XLSX file and get the first chart with Aspose.Cells | C# Aspose.Cells example to read chart name and type | Get chart collection from worksheet using Aspose.Cells .NET | Check chart count and retrieve first chart in Aspose.Cells
// Developer Intent: Load an XLSX workbook from a specified path and obtain the first chart object on the first worksheet.
// Use Cases: Log or display the name and type of the first chart for reporting purposes. | Validate that a worksheet contains at least one chart before applying further modifications. | Fetch the first chart to programmatically change its properties, such as chart type or data source.
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook from a file path, checks for charts on the first worksheet, and returns the first Chart object, handling the case where no charts exist. | Provide an Aspose.Cells snippet that iterates through all charts in a worksheet and prints each chart's Name, Type, and parent worksheet. | Create a method that accepts a file path, loads the workbook, and returns a tuple of the first chart's Name and Type, or null if the worksheet has no charts.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDemo
{
    // Demonstrates how to open an existing XLSX file with the Workbook(string) constructor, access the first worksheet, obtain its ChartCollection, verify that a chart exists, and fetch the chart at index 0. The sample prints the chart's Name, Type, and the parent worksheet name.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing XLSX file
            string filePath = "input.xlsx";

            // Load the workbook from the file path using the string constructor
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the chart collection of the worksheet
            ChartCollection charts = worksheet.Charts;

            // Ensure there is at least one chart
            if (charts.Count > 0)
            {
                // Obtain the first chart (index 0)
                Chart firstChart = charts[0];

                // Example: display some properties of the chart
                Console.WriteLine("First chart name: " + firstChart.Name);
                Console.WriteLine("First chart type: " + firstChart.Type);
                Console.WriteLine("Chart belongs to worksheet: " + firstChart.Worksheet.Name);
            }
            else
            {
                Console.WriteLine("No charts found in the first worksheet.");
            }
        }
    }
}
