// Title: Enable automatic major and minor units on a chart's Y‑axis in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Configure the ValueAxis of an Aspose.Cells chart to use automatic major and minor units in C#. | Programmatically set a chart's Y‑axis to auto‑scale major and minor units before saving the workbook with Aspose.Cells. | Apply automatic scaling to the value axis of the first chart in a worksheet using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set chart Y axis automatic major unit | How to make chart value axis auto adjust minor units with Aspose.Cells .NET | Enable auto scaling for Y‑axis in Excel chart using Aspose.Cells library | C# Aspose.Cells chart axis automatic units example
// Tags: Aspose.Cells chart valueaxis auto major unit | Aspose.Cells chart valueaxis auto minor unit | C# Excel chart axis automatic scaling Aspose.Cells | set chart Y axis automatic units .NET | auto scaling chart axis Aspose.Cells example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;   // Required for Chart class

namespace AsposeCellsExample
{
    // The sample loads an existing Excel workbook, accesses the first worksheet and its first chart, enables automatic major and minor units on the chart's value (Y) axis, and saves the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
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

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the first worksheet.");
                    return;
                }

                // Get the first chart
                Chart chart = sheet.Charts[0];

                // Enable automatic major/minor units on the Y‑axis (value axis)
                chart.ValueAxis.IsAutomaticMajorUnit = true;
                chart.ValueAxis.IsAutomaticMinorUnit = true;

                // Save the modified workbook
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
