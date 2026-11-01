// Title: Set the vertical axis minimum to 0 and maximum to 200 for the first chart in an existing Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel file with Aspose.Cells in C#, turn off automatic scaling on the chart's ValueAxis, assign MinValue = 0 and MaxValue = 200, and save the workbook. | Programmatically fix the Y‑axis range of the first chart in a .xlsx workbook to 0‑200 by setting chart.ValueAxis.IsAutomaticMinValue/IsAutomaticMaxValue to false and specifying MinValue and MaxValue. | Using Aspose.Cells, locate the first worksheet's first chart, disable its automatic axis limits, set explicit minimum and maximum values, and write the changes back to a new file.
// Common Searches: Aspose.Cells C# set chart vertical axis minimum and maximum values | how to fix chart value axis range to 0-200 with Aspose.Cells .NET | disable automatic axis scaling for Excel chart using Aspose.Cells library | change Y‑axis limits of the first chart in an existing workbook programmatically
// Tags: Aspose.Cells chart valueaxis minmax | C# set chart vertical axis range | Aspose.Cells disable automatic axis scaling | modify existing workbook chart axis .NET | Excel chart fixed Y-axis Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example loads an existing Excel workbook, accesses the first worksheet's first chart, disables automatic scaling on its value (vertical) axis, sets the axis minimum to 0 and maximum to 200, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the first worksheet.");
                return;
            }

            // Get the first chart on the worksheet
            Chart chart = sheet.Charts[0];

            try
            {
                // Disable automatic scaling for the vertical (value) axis
                chart.ValueAxis.IsAutomaticMaxValue = false;
                chart.ValueAxis.IsAutomaticMinValue = false;

                // Set the vertical axis maximum to 200 and minimum to 0
                chart.ValueAxis.MaxValue = 200;
                chart.ValueAxis.MinValue = 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error configuring chart axis: {ex.Message}");
                return;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
