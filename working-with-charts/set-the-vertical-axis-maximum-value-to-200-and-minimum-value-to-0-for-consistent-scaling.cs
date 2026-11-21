// Title: How to set the vertical (value) axis minimum to 0 and maximum to 200 for an Excel chart using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart Y-axis minimum and maximum values | programmatically define vertical axis range for Excel chart with Aspose.Cells | C# change value axis scaling of a chart in an existing workbook using Aspose.Cells | set fixed Y-axis limits 0 to 200 in Aspose.Cells generated Excel chart | adjust chart axis bounds in .NET Excel file with Aspose.Cells API
// Tags: Aspose.Cells chart value axis range | set vertical axis maxvalue Aspose.Cells | define chart Y-axis minvalue .NET | Excel chart axis scaling C# | modify chart axis bounds Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel workbook, verifies that a chart exists on the first worksheet, accesses the chart's value (vertical) axis, sets its MaxValue to 200 and MinValue to 0, and saves the modified workbook to a new file.
class Program
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count > 0)
            {
                // Retrieve the first chart
                Chart chart = sheet.Charts[0];

                // Access the vertical (value) axis
                Axis verticalAxis = chart.ValueAxis; // equivalent to chart.Axis[0]

                // Set the desired range directly
                verticalAxis.MaxValue = 200;
                verticalAxis.MinValue = 0;
            }
            else
            {
                Console.WriteLine("No charts found in the first worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
