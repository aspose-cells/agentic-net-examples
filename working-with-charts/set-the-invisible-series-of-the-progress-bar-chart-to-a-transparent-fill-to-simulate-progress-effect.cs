// Title: Make all series in an Excel progress bar chart transparent using Aspose.Cells for .NET (C#)
// AI Prompts: Load a workbook, locate the first chart, and set each series' FillFormat.Transparency to 1.0 with a solid fill using Aspose.Cells in C#. | Write C# code that iterates over a chart's NSeries collection and applies a fully transparent fill to hide the series in a progress bar chart. | Update an Excel file so that the progress bar chart appears empty by making its series area fill completely transparent via the Aspose.Cells API.
// Common Searches: Aspose.Cells C# set chart series fill transparency to 100% | how to hide series in an Excel progress bar chart using Aspose.Cells | make chart series invisible Aspose.Cells .NET | transparent fill for chart series Aspose.Cells example | C# code to set series area fill type solid and transparency 1.0 in Excel chart
// Tags: Aspose.Cells chart series transparent fill | C# set series fill transparency Aspose.Cells | Excel progress bar chart invisible series | Aspose.Cells FillFormat.Transparency example | modify chart series fill type .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Loads an existing Excel workbook, accesses the first chart, iterates through each series, sets the series area fill to a solid type with 100 % transparency, and saves the updated workbook.
class ProgressBarChart
{
    static void Main()
    {
        string inputPath = "ProgressBar.xlsx";
        string outputPath = "ProgressBar_Updated.xlsx";

        try
        {
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

            // Ensure there is at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Assume the progress bar chart is the first chart on the sheet
            Chart chart = sheet.Charts[0];

            // Iterate through all series in the chart
            foreach (Series series in chart.NSeries)
            {
                try
                {
                    // If the series area exists, make its fill fully transparent
                    if (series.Area != null && series.Area.FillFormat != null)
                    {
                        // Use the modern FillType property
                        series.Area.FillFormat.FillType = FillType.Solid;
                        // Set transparency to 100% (fully transparent)
                        series.Area.FillFormat.Transparency = 1.0;
                    }
                }
                catch (Exception exSeries)
                {
                    Console.WriteLine($"Warning: Could not modify series. {exSeries.Message}");
                }
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
