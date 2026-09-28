// Title: Update the first chart's title to the current date in an existing Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens a .xlsx file, accesses the first worksheet's first chart, and sets its Title.Text to DateTime.Now formatted as yyyy‑MM‑dd, then saves the workbook. | Show how to safely check for the presence of a chart in a worksheet before assigning the current system date to the chart's title using Aspose.Cells for .NET. | Provide a snippet that formats the current date for a chart title, updates the label, and includes error handling for missing input files or empty chart collections.
// Common Searches: aspnet aspocells set chart title to today’s date programmatically | C# Aspose.Cells change Excel chart label to current date | how to update Excel chart title with DateTime.Now using Aspose.Cells | example code for modifying chart title in existing workbook Aspose.Cells .NET | check for charts before updating title Aspose.Cells C#
// Tags: chart title update Aspose.Cells | set Excel chart label date .NET | modify existing chart title C# | Aspose.Cells chart title formatting | handle missing chart collection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // Loads an existing .xlsx workbook, retrieves the first chart on the first worksheet, assigns the current date (formatted yyyy‑MM‑dd) to the chart's Title.Text, and saves the file, with checks for missing files and empty chart collections.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                    return;
                }

                // Retrieve the first chart on the worksheet
                Chart chart = worksheet.Charts[0];

                // Update the chart's title to display the current date
                chart.Title.Text = DateTime.Now.ToString("yyyy-MM-dd");

                // Save the workbook with the updated chart title
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
