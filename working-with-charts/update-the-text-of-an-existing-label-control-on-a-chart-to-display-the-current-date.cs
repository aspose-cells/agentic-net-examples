// Title: Set an Excel chart title to today's date programmatically with Aspose.Cells for .NET (C#)
// Common Searches: C# code to change an Excel chart title to the current date with Aspose.Cells | Aspose.Cells .NET update chart title text dynamically | Set Excel chart label to today’s date using Aspose.Cells library | Programmatically modify chart title in an existing workbook with Aspose.Cells | How to format chart title as yyyy-MM-dd in Aspose.Cells for .NET
// Tags: Aspose.Cells set chart title text | Excel chart title dynamic date C# | update chart label programmatically Aspose.Cells | modify chart title in workbook .NET | chart title formatting Aspose.Cells | C# Aspose.Cells chart title automation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing workbook, accesses the first chart, sets its Title.Text to the current date formatted as yyyy-MM-dd, and saves the file, with file‑existence checks and exception handling.
class UpdateChartLabel
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

            // Load the existing workbook that contains the chart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Access the first chart on the worksheet (adjust index if needed)
            Chart chart = sheet.Charts[0];

            // Update the chart title (label) to display the current date
            chart.Title.Text = DateTime.Now.ToString("yyyy-MM-dd");

            // Save the workbook with the updated chart label
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
