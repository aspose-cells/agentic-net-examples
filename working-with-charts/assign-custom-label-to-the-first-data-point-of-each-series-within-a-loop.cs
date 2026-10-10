// Title: Assign a custom label to the first data point of each series in an Excel chart with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code using Aspose.Cells that iterates over all series in the first chart of a worksheet and sets the Series.Name property to a custom label for the first data point. | Show how to load or create an Excel workbook, locate its first chart, and apply a unique label to the first point of each series using the Aspose.Cells .NET API.
// Common Searches: Aspose.Cells C# set series name for first point in Excel chart | loop through chart series and assign custom label to first data point using Aspose.Cells | how to change the label of the first data point in each series of an Excel chart in .NET | C# Aspose.Cells example modify chart series name based on first data point | programmatically label first point of each series in Excel chart with Aspose.Cells
// Tags: set series name first data point Aspose.Cells | iterate chart series C# Aspose.Cells | modify chart series label .NET | Excel chart series labeling Aspose.Cells | custom label first point Excel chart .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The program loads (or creates) an Excel workbook, accesses the first worksheet's first chart, iterates through each series, assigns a custom name representing the first data point, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Load existing workbook if present; otherwise create a new one.
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Ensure there is at least one worksheet.
                Worksheet sheet;
                if (workbook.Worksheets.Count > 0)
                {
                    sheet = workbook.Worksheets[0];
                }
                else
                {
                    int addedIndex = workbook.Worksheets.Add();
                    sheet = workbook.Worksheets[addedIndex];
                }

                // Ensure there is at least one chart on the worksheet.
                Chart chart = sheet.Charts.Count > 0 ? sheet.Charts[0] : null;
                if (chart == null)
                {
                    Console.WriteLine("No chart found on the first worksheet. Operation skipped.");
                }
                else
                {
                    // Iterate through each series in the chart.
                    for (int i = 0; i < chart.NSeries.Count; i++)
                    {
                        Series series = chart.NSeries[i];

                        // As the Series.DataPoints property may not be available in older versions,
                        // we set a custom name for the series instead.
                        series.Name = $"Series {i + 1} - First Point";
                    }
                }

                // Save the workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
