// Title: Set scientific notation (0.00E+00) for data labels of the fourth series in an Excel chart using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# how to format chart series data labels in scientific notation | set number format 0.00E+00 for the fourth series data labels in an Excel chart with Aspose.Cells | enable data labels for a specific series and apply custom number format using Aspose.Cells | C# example to apply scientific notation to high‑magnitude chart values with Aspose.Cells | change data label format of series index 3 in an Aspose.Cells chart
// Tags: chart series data label number format Aspose.Cells | apply scientific notation to chart data labels C# | set data label format for specific series Aspose.Cells | Excel chart fourth series formatting Aspose.Cells | custom number format 0.00E+00 Aspose.Cells chart

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample loads an existing workbook, accesses the first worksheet and its first chart, enables data labels for the fourth series (index 3), applies the scientific notation format "0.00E+00" to those labels, and saves the workbook while handling missing files, absent charts, and out‑of‑range series.
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
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("Error: No charts found on the first worksheet.");
                return;
            }

            // Access the first chart on the sheet
            Chart chart = sheet.Charts[0];

            // Index of the fourth series (zero‑based)
            int seriesIndex = 3;

            // Verify that the requested series exists
            if (chart.NSeries.Count > seriesIndex)
            {
                // Enable data labels for this series
                // (Setting ShowValue to true displays the values as data labels)
                chart.NSeries[seriesIndex].DataLabels.ShowValue = true;

                // Apply scientific notation number format to the data labels
                chart.NSeries[seriesIndex].DataLabels.NumberFormat = "0.00E+00";
            }
            else
            {
                Console.WriteLine($"Warning: Series index {seriesIndex} is out of range. Series count: {chart.NSeries.Count}");
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
