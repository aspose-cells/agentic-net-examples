// Title: Add value and category data labels to the first chart in an Excel workbook while preserving the original language with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, verifies a chart exists, enables ShowValue and ShowCategoryName for each series of the first chart, hides the series name, and saves the workbook. | Show how to add data labels to a localized chart using Aspose.Cells, creating the output folder if it does not already exist. | Provide a snippet that checks the input file, accesses the first worksheet’s chart collection, configures data labels without changing the chart’s language, and writes the result to a new file.
// Common Searches: how to enable value and category data labels on an existing Excel chart with Aspose.Cells C# | preserve localized chart labels when adding data labels using Aspose.Cells for .NET | C# Aspose.Cells add data labels to first chart in workbook | check chart existence before updating data labels Aspose.Cells example
// Tags: add data labels to Aspose.Cells chart series | show value and category names in Excel chart C# | preserve original language when updating chart labels | verify chart presence before modifying Aspose.Cells workbook | ensure output directory exists when saving workbook

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Loads an existing Excel file, accesses the first worksheet's first chart, enables value and category data labels for each series while hiding the series name to keep the original language, creates the output folder if needed, and saves the modified workbook.
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

            // Load the workbook that contains the chart
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Retrieve the first chart on the sheet
            Chart chart = sheet.Charts[0];

            // Add data labels to each series
            foreach (Series series in chart.NSeries)
            {
                // Enable data labels and configure visibility
                series.DataLabels.ShowValue = true;               // Show the data point value
                series.DataLabels.ShowCategoryName = true;        // Show the category (X‑axis) name
                series.DataLabels.ShowSeriesName = false;         // Do not show series name (preserves original language)
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
