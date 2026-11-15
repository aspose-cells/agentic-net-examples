// Title: Set a custom currency number format for data labels in the first series of an Excel chart using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set currency number format for chart data labels | How to apply a custom number format to the first series data labels in an Excel chart using Aspose.Cells | C# code to format chart series data labels as $#,##0.00 with Aspose.Cells | Changing data label format to currency in Aspose.Cells chart series | Apply custom number format to Excel chart data labels with Aspose.Cells .NET
// Tags: chart series data label currency formatting Aspose.Cells | set data label number format C# Aspose.Cells | first series data label formatting Aspose.Cells | Aspose.Cells chart data label number format customization | apply currency pattern to Excel chart data labels with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, verifies that the first worksheet contains a chart with at least one series, enables data labels for the first series, assigns the currency pattern "$#,##0.00" to the data labels' NumberFormat property, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("Error: No charts found in the first worksheet.");
                    return;
                }

                // Get the first chart on the worksheet
                Chart chart = worksheet.Charts[0];

                // Ensure the chart has at least one series
                if (chart.NSeries.Count == 0)
                {
                    Console.WriteLine("Error: The chart does not contain any series.");
                    return;
                }

                // Get the first series of the chart
                var series = chart.NSeries[0];

                // Ensure data labels are displayed
                series.DataLabels.ShowValue = true;

                // Apply a custom currency number format to all data labels in this series
                series.DataLabels.NumberFormat = "$#,##0.00";

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
