// Title: How to programmatically change the title of a "Progress Bar" chart in an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set chart title based on project phase variable | Update Excel chart title dynamically with Aspose.Cells | Locate a specific chart by name and modify its title in Aspose.Cells | Make chart title visible and change its text using Aspose.Cells C# | Load workbook, find "Progress Bar" chart and rename its title programmatically
// Tags: Aspose.Cells chart title modification | set chart title text Aspose.Cells C# | locate chart by name Aspose.Cells | dynamic Excel chart title .NET | progress bar chart title update Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example loads an existing Excel workbook, searches the first worksheet for a chart named "Progress Bar", makes the chart title visible, sets the title text to include the current project phase, and saves the workbook to a new file while handling missing files and other exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Current project phase – this could come from any source (e.g., database, UI)
            string currentPhase = "Design Phase";

            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume the chart is on the first worksheet; adjust index if needed
            Worksheet sheet = workbook.Worksheets[0];

            // Locate the chart named "Progress Bar"
            Chart progressChart = null;
            foreach (Chart chart in sheet.Charts)
            {
                if (chart.Name == "Progress Bar")
                {
                    progressChart = chart;
                    break;
                }
            }

            if (progressChart != null)
            {
                try
                {
                    // Ensure the chart title is visible and set its text dynamically
                    progressChart.Title.IsVisible = true;
                    progressChart.Title.Text = $"Progress Bar - {currentPhase}";
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error updating chart title: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("Chart 'Progress Bar' not found.");
            }

            // Save the workbook with the updated chart title
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
