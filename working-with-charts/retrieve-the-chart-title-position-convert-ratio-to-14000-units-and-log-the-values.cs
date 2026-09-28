// Title: Retrieve chart title X/Y position ratios from an Excel workbook using Aspose.Cells for .NET and convert them to 1/4000 units
// AI Prompts: Generate C# code with Aspose.Cells that opens an .xlsx file, reads the X and Y position ratios of each chart title, converts the ratios to 1/4000 units, and prints the results to the console. | Enhance the example to return a dictionary of chart names mapped to their title coordinates in 1/4000 units, skipping any chart that lacks a title. | Add robust error handling that logs a warning when a chart has no title and continues processing the remaining charts.
// Common Searches: Aspose.Cells get chart title X Y position ratios in C# | convert chart title position ratio to 1/4000 units using .NET | C# read chart title coordinates from Excel workbook with Aspose.Cells | handle missing chart title when retrieving title position in Aspose.Cells
// Tags: Aspose.Cells chart title coordinates extraction | convert title ratio to 1/4000 units C# | read chart title position Aspose.Cells | missing chart title handling Aspose.Cells | iterate worksheet charts Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an Excel workbook, accesses the first worksheet's charts, reads each chart title's X and Y position ratios, converts the ratios to integer values in 1/4000 units, and writes the converted coordinates to the console, with optional handling for charts without titles.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the first worksheet.");
                return;
            }

            // Get the first chart on the worksheet
            Chart chart = sheet.Charts[0];

            // Retrieve the title position ratios (values between 0 and 1)
            // Title.X and Title.Y represent the position ratios.
            double ratioX = chart.Title.X; // Horizontal position ratio
            double ratioY = chart.Title.Y; // Vertical position ratio

            // Convert ratios to 1/4000 units
            int posX = (int)Math.Round(ratioX * 4000);
            int posY = (int)Math.Round(ratioY * 4000);

            // Log the converted positions
            Console.WriteLine($"Chart Title Position X: {posX} /4000 units");
            Console.WriteLine($"Chart Title Position Y: {posY} /4000 units");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors (e.g., loading issues, Aspose.Cells exceptions)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
