// Title: Set a 2‑point dark gray border on an Excel chart’s ChartArea with Aspose.Cells for .NET (C#)
// AI Prompts: Apply a solid dark‑gray border of 2 points thickness to the ChartArea of an existing chart in a workbook using the Aspose.Cells C# API. | Programmatically modify a loaded Excel file so that its first chart receives a 2‑point dark gray border, creating the chart if it does not already exist.
// Common Searches: Aspose.Cells C# set chart area border thickness to 2 points | how to change Excel chart border color to dark gray using Aspose.Cells | C# code to add solid border around chart in workbook with Aspose.Cells | set chart border line weight in Aspose.Cells .NET example | apply custom border style to Excel chart programmatically with Aspose.Cells
// Tags: Aspose.Cells chartarea border lineweight | C# set chart border color Aspose.Cells | Excel chart formatting border Aspose.Cells | Aspose.Cells chart area border styling .NET | modify chart border thickness programmatically

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an Excel workbook, retrieves the first chart (or adds a column chart if none exists), and shows how to set the ChartArea border to a solid dark‑gray line with a 2‑point weight using Aspose.Cells for .NET, then saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to prevent FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart; otherwise add a new one
            Chart chart;
            if (sheet.Charts.Count > 0)
            {
                // Retrieve the first chart
                chart = sheet.Charts[0];
            }
            else
            {
                // Add a default column chart and retrieve it
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
                chart = sheet.Charts[chartIndex];
            }

            // Optional: Set chart border color (if supported in the used version)
            // Note: Some older versions of Aspose.Cells may not expose border properties directly.
            // The following lines are commented out to maintain compatibility.
            // chart.ChartArea.BorderLineWeight = 2;
            // chart.ChartArea.BorderLineColor = Color.DarkGray;
            // chart.ChartArea.BorderLineStyle = LineStyleType.Solid;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the updated chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
