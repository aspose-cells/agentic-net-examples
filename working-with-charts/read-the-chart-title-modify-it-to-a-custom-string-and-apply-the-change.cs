// Title: Read the first chart title and set a custom title in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing .xlsx workbook, retrieve the text of the first chart's title, replace it with a custom string, ensure the title is visible, and save the file using Aspose.Cells in C#. | Show how to verify a chart exists, read its current title, modify the Title.Text and Title.IsVisible properties, and write the updated workbook back to disk with Aspose.Cells.
// Common Searches: C# Aspose.Cells read chart title from first worksheet | how to change Excel chart title programmatically using Aspose.Cells .NET | set custom chart title and make it visible with Aspose.Cells in C# | Aspose.Cells example for updating chart title and saving workbook | check if chart exists before modifying title Aspose.Cells C#
// Tags: Aspose.Cells modify chart title C# | update Excel chart title .NET | read chart title Aspose.Cells | chart title visibility Aspose.Cells | first chart manipulation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads input.xlsx, verifies a chart exists, reads the first chart's title, replaces it with "Sales Performance Q1 2024", sets Title.IsVisible = true, and saves the workbook as output.xlsx.
class ModifyChartTitle
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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart in the worksheet
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart
            Chart chart = sheet.Charts[0];

            // Read the current chart title
            string originalTitle = chart.Title.Text;
            Console.WriteLine("Original Chart Title: " + originalTitle);

            // Modify the chart title to a custom string
            string customTitle = "Sales Performance Q1 2024";
            chart.Title.Text = customTitle;

            // Ensure the title is visible
            chart.Title.IsVisible = true;

            Console.WriteLine("Modified Chart Title: " + chart.Title.Text);

            // Save the workbook with the updated chart title
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
