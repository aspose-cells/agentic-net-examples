// Title: How to read category and value axis labels from an Aspose.Cells chart after calling Chart.Calculate in C#
// Common Searches: Aspose.Cells C# get chart axis labels after Chart.Calculate | retrieve category axis text from an Excel chart using Aspose.Cells | read value axis tick labels from a calculated chart in .NET | C# code to extract chart axis texts after updating chart data with Aspose.Cells | how to use GetAxisTexts with Aspose.Cells chart objects
// Tags: chart.Calculate axis label extraction | CategoryAxis.GetAxisTexts usage | ValueAxis.GetAxisTexts usage | Aspose.Cells chart axis text retrieval | C# read Excel chart axis labels

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example loads an Excel workbook, accesses the first worksheet, ensures a chart exists, calls Chart.Calculate() to refresh the chart data, then obtains the category and value axis labels via CategoryAxis.GetAxisTexts() and ValueAxis.GetAxisTexts(), prints them to the console, and optionally saves the workbook.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Retrieve the first chart
            Chart chart = sheet.Charts[0];

            // Calculate the chart to update axis values
            chart.Calculate();

            // Get axis labels
            string[] categoryAxisLabels = chart.CategoryAxis.GetAxisTexts();
            string[] valueAxisLabels = chart.ValueAxis.GetAxisTexts();

            // Output the axis labels
            Console.WriteLine("Category Axis Labels:");
            foreach (string label in categoryAxisLabels)
            {
                Console.WriteLine(label);
            }

            Console.WriteLine("Value Axis Labels:");
            foreach (string label in valueAxisLabels)
            {
                Console.WriteLine(label);
            }

            // Save the workbook (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
