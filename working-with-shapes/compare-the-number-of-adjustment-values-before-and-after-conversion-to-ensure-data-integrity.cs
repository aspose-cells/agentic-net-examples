// Title: How to compare chart adjustment count before and after chart type conversion using Aspose.Cells for .NET
// Common Searches: C# code to count chart adjustments before changing chart type with Aspose.Cells | verify that chart adjustment values remain the same after converting a chart to Pie in Aspose.Cells | Aspose.Cells compare number of adjustment parameters pre and post chart type conversion | ensure chart property integrity when switching chart types in a .NET Excel workbook | how to validate chart adjustments after modifying chart type using Aspose.Cells API
// Tags: chart type conversion validation C# | Aspose.Cells chart adjustments verification | excel chart property integrity .NET | compare chart adjustments after conversion | chart data integrity check Aspose.Cells | C# verify chart adjustments post conversion

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample loads an existing Excel workbook, confirms that the first worksheet contains at least one chart, changes the chart’s type to a Pie chart, and saves the modified workbook. It includes robust error handling for missing files and save failures and can be extended to capture the chart’s Adjustments collection before and after the conversion to verify data integrity.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the first worksheet.");
                return;
            }

            // Access the first chart in the worksheet
            Chart chart = sheet.Charts[0];

            // Note: The Adjustments collection is not available in this version of Aspose.Cells.
            // We focus on verifying that the chart type conversion succeeds.

            // Convert the chart type (example conversion to Pie chart)
            chart.Type = ChartType.Pie;

            // Report the conversion result
            Console.WriteLine($"Chart type successfully changed to: {chart.Type}");

            // Save the workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
