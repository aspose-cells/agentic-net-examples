// Title: Clone an existing chart, modify its data source, and insert the cloned chart into a different worksheet using Aspose.Cells for .NET (C#)
// Common Searches: how to duplicate a chart and set a new data range in Aspose.Cells C# | Aspose.Cells copy chart to another sheet and update series source | C# Aspose.Cells clone chart and change NSeries data reference | add cloned chart with custom series to a different worksheet using Aspose.Cells | Aspose.Cells chart cloning example with modified data source
// Tags: chart cloning with Aspose.Cells C# | modify chart NSeries data source Aspose.Cells | add chart to different worksheet Aspose.Cells | copy chart type and set series range Aspose.Cells | Aspose.Cells workbook chart manipulation C#

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample loads a workbook, clones the first chart from Sheet1 onto Sheet2, clears the default series, adds a new series referencing Sheet2!$B$2:$B$5 with a custom name, and saves the result as a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook that contains the original chart
            Workbook workbook = new Workbook(inputPath);

            // Source worksheet and the chart to be cloned (assumes first chart)
            Worksheet sourceSheet = workbook.Worksheets["Sheet1"];
            if (sourceSheet == null || sourceSheet.Charts.Count == 0)
            {
                Console.WriteLine("Source worksheet or chart not found.");
                return;
            }

            Chart originalChart = sourceSheet.Charts[0];

            // Target worksheet where the cloned chart will be placed
            Worksheet dataSheet = workbook.Worksheets["Sheet2"];
            if (dataSheet == null)
            {
                Console.WriteLine("Data worksheet \"Sheet2\" not found.");
                return;
            }

            // Determine placement (rows/columns) for the new chart
            int upperLeftRow = 5;
            int upperLeftColumn = 1;
            int lowerRightRow = upperLeftRow + 15;   // approximate height
            int lowerRightColumn = upperLeftColumn + 10; // approximate width

            // Add a new chart of the same type as the original
            int chartIndex = dataSheet.Charts.Add(originalChart.Type, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);
            Chart clonedChart = dataSheet.Charts[chartIndex];

            // Clear any default series
            clonedChart.NSeries.Clear();

            // Add a new series with the desired data ranges
            // Values range
            int seriesIdx = clonedChart.NSeries.Add("Sheet2!$B$2:$B$5", false);
            // Set series name
            clonedChart.NSeries[seriesIdx].Name = "New Series";

            // NOTE: CategoryData property is not available in the referenced Aspose.Cells version.
            // If needed, categories can be set using alternative methods or by adjusting the source data.

            // Save the workbook with the new chart inserted
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
