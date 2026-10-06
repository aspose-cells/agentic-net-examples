// Title: How to dynamically set the title of a "Progress Bar" chart in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that loads or creates an Excel workbook, locates a chart named "Progress Bar", makes its title visible, and assigns a formatted string containing a project phase variable. | Show how to programmatically add a placeholder column chart named "Progress Bar" when it is missing, then update its title text to reflect the current project phase using Aspose.Cells. | Demonstrate updating an Excel chart title at runtime with Aspose.Cells, including error handling for missing files and saving the modified workbook.
// Common Searches: Aspose.Cells C# change chart title based on variable value | set dynamic title for specific chart in Excel using Aspose.Cells .NET | create or retrieve chart by name and update title in Aspose.Cells | make chart title visible and set text with Aspose.Cells C# | load existing workbook or create new one and modify chart title Aspose.Cells
// Tags: Aspose.Cells chart title update | C# dynamic Excel chart title | Aspose.Cells retrieve chart by name | create placeholder column chart Aspose.Cells | load or create workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing workbook or creates a new one, accesses the first worksheet, retrieves a chart named "Progress Bar" (creating a column chart placeholder if absent), makes the chart title visible, sets the title to "Project Phase: {currentPhase}" using a variable, and saves the workbook while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Define the current project phase (could be sourced elsewhere)
            string currentPhase = "Design";

            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                // Ensure at least one worksheet exists
                if (workbook.Worksheets.Count == 0)
                {
                    workbook.Worksheets.Add("Sheet1");
                }
            }

            // Access the first worksheet (adjust index/name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Try to retrieve the chart by name
            Chart progressBarChart = worksheet.Charts["Progress Bar"];

            // If the chart does not exist, create a placeholder chart
            if (progressBarChart == null)
            {
                // Add a column chart; Add returns the chart index
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                progressBarChart = worksheet.Charts[chartIndex];
                progressBarChart.Name = "Progress Bar";
            }

            // Make the title visible and set its text dynamically
            progressBarChart.Title.IsVisible = true;
            progressBarChart.Title.Text = $"Project Phase: {currentPhase}";

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
