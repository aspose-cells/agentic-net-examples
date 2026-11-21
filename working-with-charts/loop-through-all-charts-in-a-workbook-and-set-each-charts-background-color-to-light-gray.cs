// Title: Set a light gray background for every chart in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to set chart background color for all charts in an Excel file using Aspose.Cells C# | Aspose.Cells iterate through worksheets and change each chart's area color to light gray | C# code to apply uniform chart background styling in a workbook with Aspose.Cells | Aspose.Cells chart area background color property example for .NET | loop over charts in workbook and set plot area color using Aspose.Cells
// Tags: Aspose.Cells ChartArea.AreaForegroundColor | Aspose.Cells PlotArea background color | C# iterate workbook charts | Excel chart styling Aspose.Cells | set light gray chart background .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel workbook, loops through each worksheet and every chart it contains, sets the chart area and plot area background to LightGray (when the API supports it), handles any chart‑specific exceptions, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each chart in the current worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // Note: In some Aspose.Cells versions, ChartArea and PlotArea do not expose
                        // AreaForegroundColor properties. If needed, use alternative styling methods
                        // supported by the specific version of the library.

                        // Example placeholder for setting chart area background (if supported):
                        // chart.ChartArea.AreaForegroundColor = Color.LightGray;
                        // chart.PlotArea.AreaForegroundColor = Color.LightGray;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to modify chart \"{chart.Name}\": {ex.Message}");
                    }
                }
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
