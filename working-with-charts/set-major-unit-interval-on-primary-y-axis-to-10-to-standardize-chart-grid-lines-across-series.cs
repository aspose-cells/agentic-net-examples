// Title: Set a fixed major unit of 10 on the primary Y‑axis of an Excel chart using Aspose.Cells for .NET
// AI Prompts: Turn off automatic major unit calculation and assign a major unit value of 10 to the primary Y‑axis of a chart in a workbook with Aspose.Cells C#. | Configure the value axis of the first chart in an Excel file to use a fixed grid interval of 10 by setting IsAutomaticMajorUnit to false and MajorUnit to 10 via Aspose.Cells.
// Common Searches: Aspose.Cells C# set chart primary Y axis major unit to 10 | how to disable automatic major unit on Excel chart axis using Aspose.Cells | fix grid line spacing on value axis in Aspose.Cells chart | set fixed major unit for Y axis in .NET Excel chart with Aspose.Cells | customize chart axis scaling programmatically with Aspose.Cells
// Tags: Aspose.Cells chart primary Y axis major unit | disable automatic major unit Aspose.Cells | set fixed major unit Excel chart .NET | value axis scaling Aspose.Cells C# | chart grid interval customization Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Loads an Excel workbook, accesses the first worksheet's first chart, disables automatic major unit calculation on the primary Y‑axis, sets the major unit to 10, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains the chart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet has at least one chart
            if (worksheet.Charts.Count > 0)
            {
                // Get the first chart on the worksheet
                Chart chart = worksheet.Charts[0];

                try
                {
                    // Access the primary Y axis (value axis) of the chart.
                    // For older Aspose.Cells versions the property is ValueAxis.
                    Axis primaryYAxis = chart.ValueAxis;

                    // Turn off automatic calculation of the major unit
                    primaryYAxis.IsAutomaticMajorUnit = false;

                    // Set the major unit interval to 10 to standardize grid lines
                    primaryYAxis.MajorUnit = 10;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to modify chart axis: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("No charts found in the worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
