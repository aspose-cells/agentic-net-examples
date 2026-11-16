// Title: Adjust GapWidth to Thicken Bars in a Progress Bar Chart using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart GapWidth to make progress bar columns wider | How to increase bar thickness of an Excel progress bar chart with Aspose.Cells .NET | Programmatically change column chart spacing in a workbook using Aspose.Cells | C# code to reduce gap width of a chart series for thicker bars in Excel | Modify Excel chart bar width with Aspose.Cells API in .NET
// Tags: chart GapWidth property Aspose.Cells | thicker column bars Excel Aspose.Cells | reduce chart series spacing .NET | progress bar chart bar width C# | adjust Excel chart bar thickness programmatically

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example loads an existing workbook, finds a chart named "ProgressBar" (or falls back to the first chart), sets its GapWidth property to 20% to produce noticeably thicker bars, and saves the workbook as a new file.
class Program
{
    static void Main()
    {
        const string inputFile = "ProgressBarChart.xlsx";
        const string outputFile = "ProgressBarChart_ThickBars.xlsx";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file \"{inputFile}\" not found.");
                return;
            }

            // Load the workbook that contains the Progress Bar chart
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Locate the chart named "ProgressBar". If not found, fall back to the first chart.
            Chart progressChart = null;
            foreach (Chart chart in worksheet.Charts)
            {
                if (chart.Name == "ProgressBar")
                {
                    progressChart = chart;
                    break;
                }
            }

            if (progressChart == null && worksheet.Charts.Count > 0)
            {
                progressChart = worksheet.Charts[0];
            }

            // If the chart is found, adjust its gap width to make the bars thicker.
            // GapWidth is expressed as a percentage (0‑500). Smaller values produce thicker bars.
            if (progressChart != null)
            {
                progressChart.GapWidth = 20; // Example: 20% gap width
            }
            else
            {
                Console.WriteLine("No chart found in the worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved as \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
