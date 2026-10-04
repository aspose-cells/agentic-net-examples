// Title: Read formatted X and Y axis labels from an Excel chart after calling Chart.Calculate() with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx workbook using Aspose.Cells, invokes Chart.Calculate() on the first chart, and prints the category and value axis texts via Axis.GetAxisTexts(). | Show how to capture the formatted axis label strings from a chart after calculation and store them in string arrays for further processing. | Adapt the example to export the retrieved X‑axis and Y‑axis labels to a CSV file instead of writing them to the console.
// Common Searches: Aspose.Cells C# get chart axis labels after Chart.Calculate | How to retrieve formatted axis texts from an Excel chart using Aspose.Cells | Example of using Axis.GetAxisTexts() in .NET | Read X and Y axis values from a chart in an .xlsx workbook with Aspose.Cells | C# chart calculation before extracting axis labels Aspose.Cells
// Tags: Chart.Calculate with Aspose.Cells | Axis.GetAxisTexts C# | extract formatted axis labels from Excel chart | read chart axis texts after calculation | Aspose.Cells chart axis retrieval | C# Excel chart label extraction

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an Excel workbook, selects the first worksheet and its first chart, runs Chart.Calculate() to update the chart, then obtains the formatted category (X) and value (Y) axis labels via Axis.GetAxisTexts() and writes them to the console.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Ensure the workbook contains at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("Error: The workbook does not contain any worksheets.");
                return;
            }

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("Error: The worksheet does not contain any charts.");
                return;
            }

            // Get the first chart on the worksheet (adjust index if needed)
            Chart chart = worksheet.Charts[0];

            // Perform chart calculations to ensure axis data is up‑to‑date
            chart.Calculate();

            // Retrieve the category (X) axis and value (Y) axis
            Axis xAxis = chart.CategoryAxis;
            Axis yAxis = chart.ValueAxis;

            // Get the formatted axis labels after calculation
            string[] xAxisLabels = xAxis.GetAxisTexts();
            string[] yAxisLabels = yAxis.GetAxisTexts();

            // Output the axis labels
            Console.WriteLine("X Axis Labels:");
            foreach (string label in xAxisLabels)
            {
                Console.WriteLine(label);
            }

            Console.WriteLine("\nY Axis Labels:");
            foreach (string label in yAxisLabels)
            {
                Console.WriteLine(label);
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
