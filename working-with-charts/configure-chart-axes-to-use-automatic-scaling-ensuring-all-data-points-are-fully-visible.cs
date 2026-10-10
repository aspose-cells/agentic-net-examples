// Title: How to enable automatic scaling for both value and category axes of an Excel chart using Aspose.Cells in C#
// AI Prompts: Write C# code that opens an existing .xlsx workbook with Aspose.Cells, retrieves the first chart, and sets IsAutomaticMaxValue, IsAutomaticMinValue, IsAutomaticMajorUnit, and IsAutomaticMinorUnit to true for both the ValueAxis and CategoryAxis before saving the file. | Show how to configure Aspose.Cells chart axes to use automatic scaling for max, min, major and minor units in a .NET application.
// Common Searches: Aspose.Cells C# set chart value axis automatic max and min | How to make Excel chart axes auto scale using Aspose.Cells .NET | C# example for enabling automatic major and minor units on chart category axis with Aspose.Cells | Automatically adjust chart axes range in an existing workbook with Aspose.Cells
// Tags: Aspose.Cells chart axis automatic range | C# set chart value axis auto max/min | Aspose.Cells category axis automatic major unit | update existing workbook chart axes Aspose.Cells | Aspose.Cells chart axis property configuration .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an existing Excel workbook, verifies a chart on the first worksheet, enables automatic maximum, minimum, major and minor units for both the value (vertical) and category (horizontal) axes, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure that at least one chart exists on the worksheet
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            // Access the first chart
            Chart chart = sheet.Charts[0];

            // Enable automatic scaling for the value (vertical) axis
            chart.ValueAxis.IsAutomaticMaxValue = true;
            chart.ValueAxis.IsAutomaticMinValue = true;
            chart.ValueAxis.IsAutomaticMajorUnit = true;
            chart.ValueAxis.IsAutomaticMinorUnit = true;

            // Enable automatic scaling for the category (horizontal) axis
            chart.CategoryAxis.IsAutomaticMaxValue = true;
            chart.CategoryAxis.IsAutomaticMinValue = true;
            chart.CategoryAxis.IsAutomaticMajorUnit = true;
            chart.CategoryAxis.IsAutomaticMinorUnit = true;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
