// Title: How to detect a primary category (X) axis in an Excel chart using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx workbook with Aspose.Cells, retrieves the first chart, and uses chart.HasAxis(AxisType.Category, true) to return a boolean indicating the presence of a primary category axis. | Create a .NET console snippet that prints "Category axis present: true/false" after checking a chart's primary X‑axis using Aspose.Cells' HasAxis method.
// Common Searches: aspnet check if Excel chart has X axis using Aspose.Cells | C# sample to verify chart axis existence with Aspose | determine if a chart includes a category axis in .xlsx file | how to use Aspose.Cells to inspect chart axes | example of checking chart axes in a .NET console app
// Tags: Aspose.Cells chart.HasAxis method usage | detect category axis in .NET Excel chart | C# chart axis presence check with Aspose | Excel chart validation using Aspose.Cells | AxisType.Category check in C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program loads an .xlsx workbook with Aspose.Cells, accesses the first worksheet's first chart, calls chart.HasAxis(AxisType.Category, true) to determine whether a primary category (X) axis exists, and prints the boolean result to the console, with basic file existence verification and error handling.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one chart in the worksheet
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart
            Chart chart = worksheet.Charts[0];

            // Check if a primary Category (X) axis exists
            bool hasCategoryAxis = chart.HasAxis(AxisType.Category, true);

            // Output the result
            Console.WriteLine($"Category axis present: {hasCategoryAxis}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
