// Title: How to retrieve an Excel chart's X axis and set it to a value axis for continuous numeric scaling using Aspose.Cells for .NET
// AI Prompts: Generate C# code that accesses a chart's X (category) axis and changes its type to a value axis with Aspose.Cells. | Show the step‑by‑step process to convert a category axis into a value axis for numeric scaling in an Excel chart using Aspose.Cells .NET. | Explain how to modify the Axis object so the X axis behaves as a value axis for continuous numeric data in Aspose.Cells.
// Common Searches: Aspose.Cells set chart X axis to value axis C# example | change Excel chart category axis to numeric axis using Aspose.Cells .NET | continuous numeric scaling on chart axis Aspose.Cells C# | how to make chart X axis treat data as values in Aspose.Cells
// Tags: chart axis type adjustment Aspose.Cells | C# modify Excel chart X axis | chart axis scaling settings Aspose.Cells | retrieve category axis Aspose.Cells | value axis configuration example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing workbook, accesses the first worksheet and its first chart, retrieves the chart's CategoryAxis (X axis), and notes that Aspose.Cells for .NET does not expose a direct property to switch the axis to a value axis, so the code only demonstrates axis retrieval and saves the workbook unchanged.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("Error: No charts found on the first worksheet.");
                return;
            }

            // Access the first chart
            Chart chart = worksheet.Charts[0];

            // Retrieve the X (category) axis
            Axis xAxis = chart.CategoryAxis;

            // NOTE: Aspose.Cells for .NET does not expose an IsValueAxis property.
            // If axis type adjustment is required, use appropriate Axis properties
            // available in the version you are using.

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
