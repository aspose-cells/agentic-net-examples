// Title: How to dynamically set the title of a "Progress Bar" chart in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that loads a workbook, locates a chart named "Progress Bar" (or falls back to the first chart), assigns a variable containing the current project phase to the chart's Title.Text, ensures the title is visible, and saves the file. | Write a helper method in C# that determines the current project phase, formats a title like "Progress – {phase}", and applies it to a chart's Title property using Aspose.Cells. | Show how to gracefully handle a missing named chart in Aspose.Cells by checking for the chart, using a fallback chart, updating its title, and persisting the changes.
// Common Searches: Aspose.Cells C# change Excel chart title based on variable value | Set dynamic title for specific chart in .NET workbook using Aspose.Cells | How to update chart title to current project phase in C# with Aspose.Cells | Fallback to first chart if named chart not found Aspose.Cells | Make chart title visible after updating with Aspose.Cells .NET
// Tags: Aspose.Cells set chart title dynamically | C# update Excel chart title Aspose.Cells | retrieve chart by name Aspose.Cells | fallback to first chart Aspose.Cells | make chart title visible Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads Project.xlsx, finds the chart named "Progress Bar" (or the first chart if the name is missing), composes a title using the current project phase, sets the chart's Title.Text, makes the title visible, and saves the updated workbook as Project_Updated.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "Project.xlsx";
                string outputPath = "Project_Updated.xlsx";

                // Ensure the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the chart named "Progress Bar"
                Chart progressChart = null;
                try
                {
                    progressChart = worksheet.Charts["Progress Bar"];
                }
                catch
                {
                    // Ignore if not found by name
                }

                // Fallback to first chart if named chart is missing
                if (progressChart == null && worksheet.Charts.Count > 0)
                {
                    progressChart = worksheet.Charts[0];
                }

                if (progressChart == null)
                {
                    Console.WriteLine("Progress Bar chart not found in the worksheet.");
                    return;
                }

                // Determine the current project phase (replace with real logic)
                string currentPhase = GetCurrentProjectPhase();

                // Set the chart title dynamically based on the current phase
                progressChart.Title.Text = $"Progress – {currentPhase}";
                progressChart.Title.IsVisible = true; // Ensure the title is displayed

                // Save the workbook with the updated chart title
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method placeholder – replace with your real implementation
        private static string GetCurrentProjectPhase()
        {
            // Example static value; in practice, retrieve this from your data source
            return "Development";
        }
    }
}
