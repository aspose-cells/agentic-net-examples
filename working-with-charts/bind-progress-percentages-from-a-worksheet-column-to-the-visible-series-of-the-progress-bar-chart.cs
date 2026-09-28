// Title: Bind a worksheet column of progress percentages to a stacked bar chart series with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an existing Excel file, adds a stacked bar chart, and links the B2:B10 range as the visible series using Aspose.Cells. | Write a method that checks for the input workbook, creates the output folder if missing, and saves the workbook after binding progress data to the chart. | Show how to set the series name and chart title for a progress bar chart created programmatically with Aspose.Cells.
// Common Searches: Aspose.Cells C# bind worksheet column to stacked bar chart series | how to create a progress bar chart from Excel data using Aspose.Cells | programmatically add chart and link data range in .NET workbook | C# example for setting chart series name in Aspose.Cells | ensure output directory exists before saving workbook Aspose.Cells
// Tags: link column range to chart series Aspose.Cells | add stacked bar progress chart C# | set chart series name Aspose.Cells | check input workbook existence Aspose.Cells | create output directory before saving workbook | load workbook and insert chart Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel workbook, adds a stacked bar chart titled "Progress" to a specified worksheet, binds the B2:B10 range as the visible series, names the series "Progress", ensures the output directory exists, and saves the modified workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Define file paths (replace with actual paths as needed)
            string inputFilePath = "Input.xlsx";
            string outputFilePath = "Output.xlsx";
            string sheetName = "Sheet1";
            string progressColumnRange = "B2:B10"; // example range for progress percentages

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFilePath))
            {
                Console.WriteLine($"Input file not found: {inputFilePath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputFilePath);
            var worksheet = workbook.Worksheets[sheetName];
            if (worksheet == null)
            {
                Console.WriteLine($"Worksheet \"{sheetName}\" not found.");
                return;
            }

            // Add a stacked bar chart to act as a progress bar
            int chartIndex = worksheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];
            chart.Title.Text = "Progress";

            // Bind the progress percentages column to the visible series
            int seriesIndex = chart.NSeries.Add(progressColumnRange, true);
            chart.NSeries[seriesIndex].Name = "Progress";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputFilePath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputFilePath);
            Console.WriteLine($"Workbook saved successfully to {outputFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
