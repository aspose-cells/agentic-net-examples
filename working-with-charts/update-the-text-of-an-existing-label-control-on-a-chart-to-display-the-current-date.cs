// Title: Update an existing Excel chart’s title to show the current date with Aspose.Cells for .NET (C#)
// AI Prompts: Load an .xlsx workbook using Aspose.Cells, get the first chart, set chart.Title.Text to DateTime.Now formatted as yyyy-MM-dd, and save the file. | Programmatically replace a chart’s title text in a loaded workbook with today’s date using C# and the Aspose.Cells API. | Use Aspose.Cells for .NET to assign the current date to a chart label and write the updated workbook back to disk.
// Common Searches: Aspose.Cells C# change chart title to today's date in existing workbook | How to set Excel chart label to current date using .NET library | Update chart title dynamically with DateTime.Now in Aspose.Cells
// Tags: Aspose.Cells set chart title | C# update Excel chart label | dynamic chart title date Aspose | modify existing chart title .NET | Excel chart title formatting with DateTime

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Loads an existing .xlsx file, accesses the first chart on the first worksheet, replaces its title with the current date (yyyy-MM-dd), and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one chart in the worksheet
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the target chart (assumes the first chart is the one to modify)
            Chart chart = worksheet.Charts[0];

            // Update the chart's title to display the current date
            chart.Title.Text = DateTime.Now.ToString("yyyy-MM-dd");

            // Save the workbook with the updated chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
