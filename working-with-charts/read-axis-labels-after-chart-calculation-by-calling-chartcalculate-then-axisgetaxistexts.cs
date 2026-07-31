// Title: Read Axis Labels from an Aspose.Cells Chart (C#) using Chart.Calculate & GetAxisTexts
// Description: Creates a workbook, adds a column chart, calls Chart.Calculate to generate layout, then extracts value and category axis labels with GetAxisTexts, prints them, and saves the file.
// Keywords: Aspose.Cells chart axis labels | Chart.Calculate C# | GetAxisTexts example | retrieve value axis texts | category axis texts Aspose.Cells | Aspose.Cells .NET chart API | read chart labels programmatically
// Common Searches: Aspose.Cells get axis labels after calculate | C# retrieve chart value axis texts | How to read category axis labels Aspose.Cells | Chart.Calculate required for GetAxisTexts | Aspose.Cells chart axis text extraction
// Developer Intent: Extract the automatically generated value and category axis labels from a chart after it has been calculated.
// Use Cases: Show calculated axis labels in a console or UI for verification | Export axis labels to a report or another file format | Validate chart labeling in automated unit tests
// AI Prompts: Provide C# code that extracts both value and category axis texts from an Aspose.Cells chart after calling Chart.Calculate. | Explain why Chart.Calculate must be invoked before using GetAxisTexts and describe the returned data format. | Show how to store the retrieved axis labels in a collection for further processing.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AxisLabelsDemo
{
    // Creates a workbook, adds a column chart, calls Chart.Calculate to generate layout, then extracts value and category axis labels with GetAxisTexts, prints them, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(8000);
            worksheet.Cells["B3"].PutValue(4000);
            worksheet.Cells["B4"].PutValue(-8000);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Calculate the chart to generate axis labels
            chart.Calculate();

            // Retrieve the axis labels after calculation
            // Example: get labels from the value axis
            string[] valueAxisLabels = chart.ValueAxis.GetAxisTexts();

            // Output the retrieved labels
            Console.WriteLine("Value Axis Labels:");
            foreach (string label in valueAxisLabels)
            {
                Console.WriteLine(label);
            }

            // Optionally, retrieve category axis labels as well
            string[] categoryAxisLabels = chart.CategoryAxis.GetAxisTexts();
            Console.WriteLine("\nCategory Axis Labels:");
            foreach (string label in categoryAxisLabels)
            {
                Console.WriteLine(label);
            }

            // Save the workbook (optional, just to demonstrate lifecycle)
            workbook.Save("AxisLabelsDemo.xlsx");
        }
    }
}
