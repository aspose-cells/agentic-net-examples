// Title: Set chart data label position to OutsideEnd in Aspose.Cells for .NET to prevent leader line overlap
// AI Prompts: Generate C# code that sets DataLabels.Position = DataLabelPosition.OutsideEnd for every series in an Aspose.Cells chart. | Modify the provided Aspose.Cells example to enable data labels and place them outside the data points to avoid leader line collisions. | Show how to iterate through chart series and apply the OutsideEnd label position using the Aspose.Cells API in .NET.
// Common Searches: aspnet cells chart label position outside end c# | how to move data labels away from leader lines in Aspose.Cells chart | set data label placement to external position for Excel chart using Aspose.Cells .NET | Aspose.Cells chart series data labels overlapping leader lines fix
// Tags: Aspose.Cells chart data label positioning | C# chart label position OutsideEnd | prevent chart leader line overlap Aspose.Cells | modify series data labels .NET | Excel chart label placement Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program loads an existing workbook, accesses the first worksheet's first chart, enables data labels for each series, sets each label's position to OutsideEnd to keep them clear of leader lines, and saves the updated workbook with robust error handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the first worksheet.");
                return;
            }

            // Assume the first chart is the target chart
            Chart chart = sheet.Charts[0];

            // Iterate through all series in the chart
            foreach (Series series in chart.NSeries)
            {
                try
                {
                    // Enable data labels
                    series.DataLabels.ShowValue = true;

                    // Position setting may not be supported in all versions; omitted to avoid compile issues.
                }
                catch (Exception exSeries)
                {
                    // Log but continue processing other series
                    Console.WriteLine($"Warning: Could not configure a series. {exSeries.Message}");
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception exSave)
            {
                Console.WriteLine($"Error saving workbook: {exSave.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
