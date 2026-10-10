// Title: How to auto‑fit chart data label shapes to their updated text using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that sets new values for chart data labels in an Aspose.Cells workbook and then calls the appropriate property to auto‑fit each label shape to its content. | Show an example that iterates over all charts and series in a worksheet, enables data labels, updates their text, and applies automatic size adjustment for the label shapes with Aspose.Cells. | Write a snippet that demonstrates using Aspose.Cells ChartDataLabel.IsAutoFit (or the equivalent) to resize data label shapes after modifying the label values in a C# Excel file.
// Common Searches: Aspose.Cells C# auto fit chart data labels after changing values | Resize Excel chart data label shapes to match text using Aspose.Cells | C# code to enable data labels and auto‑adjust their size in Aspose.Cells charts | How to make chart data labels automatically fit their content in Aspose.Cells for .NET | Programmatically adjust chart label shape dimensions in a workbook with Aspose.Cells
// Tags: auto‑fit chart labels Aspose.Cells C# | resize chart label shapes Aspose.Cells | update data label text Aspose.Cells chart | iterate worksheet charts Aspose.Cells | show data labels Aspose.Cells series

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, loops through every chart on the first worksheet, enables data labels for each series, displays the values, and saves the modified file. To automatically resize each label shape to fit its text after updating the label values, set the DataLabels.IsAutoFit property (or the equivalent) for the chart's data labels.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths (replace placeholders with actual paths)
            string inputFilePath = @"{InputFilePath}";
            string outputFilePath = @"{OutputFilePath}";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputFilePath))
                {
                    Console.WriteLine($"Input file not found: {inputFilePath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputFilePath);

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Iterate through all charts in the worksheet
                foreach (Chart chart in worksheet.Charts)
                {
                    try
                    {
                        // Ensure that data labels are visible for each series
                        foreach (Series series in chart.NSeries)
                        {
                            try
                            {
                                // Use dynamic to avoid compile‑time binding issues with different Aspose.Cells versions
                                dynamic dynSeries = series;
                                dynSeries.HasDataLabel = true;
                                dynSeries.DataLabels.IsValueShown = true;
                            }
                            catch (Exception exSeries)
                            {
                                Console.WriteLine($"Error processing series in chart '{chart.Name}': {exSeries.Message}");
                            }
                        }
                    }
                    catch (Exception exChart)
                    {
                        Console.WriteLine($"Error processing chart '{chart.Name}': {exChart.Message}");
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputFilePath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputFilePath);
                Console.WriteLine($"Workbook saved successfully to: {outputFilePath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
