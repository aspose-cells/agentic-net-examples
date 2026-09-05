// Title: Hide a specific series in a stacked bar chart with Aspose.Cells for .NET (C#)
// AI Prompts: Remove the second data series from a stacked bar chart in an existing Excel file using Aspose.Cells C#. | Set a chart series' visibility to false in a stacked bar chart and save the workbook with Aspose.Cells for .NET. | Programmatically delete a chosen series from the first chart of a workbook and export the updated file using Aspose.Cells C#.
// Common Searches: Aspose.Cells C# hide series in stacked bar chart | remove specific series from Excel chart using Aspose.Cells .NET | how to make a chart series invisible in Aspose.Cells C# | delete a data series from a stacked bar chart programmatically Aspose.Cells | Aspose.Cells hide chart series without deleting data
// Tags: chart series visibility Aspose.Cells C# | remove series from stacked bar chart Aspose.Cells | Aspose.Cells hide chart series .NET | Excel chart series manipulation Aspose.Cells | stacked bar chart series removal C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads 'input.xlsx', verifies the first worksheet contains a chart, removes the series at index 1 (effectively hiding it) from a stacked bar chart, and saves the modified workbook as 'output.xlsx', with comprehensive error handling for missing files and absent charts.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook that contains the stacked bar chart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("Error: No charts found on the first worksheet.");
                return;
            }

            // Get the first chart on the worksheet (adjust index if needed)
            Chart chart = sheet.Charts[0];

            // Hide (remove) the series you want (e.g., the second series, index 1)
            int seriesIndexToHide = 1;
            if (seriesIndexToHide < chart.NSeries.Count)
            {
                try
                {
                    // Removing the series effectively hides it from the chart
                    chart.NSeries.RemoveAt(seriesIndexToHide);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to hide series at index {seriesIndexToHide}: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Warning: Series index {seriesIndexToHide} is out of range.");
            }

            // Save the workbook with the updated chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
