// Title: Apply scientific notation number format to the fourth series data labels in an Excel chart with Aspose.Cells for .NET
// AI Prompts: Enable data labels for the fourth series (index 3) of a chart and set its NumberFormat to "0.00E+00" using Aspose.Cells in C#. | Load an existing workbook, locate the first chart, and format the fourth series' data labels to display values in scientific notation. | Modify an Excel file so that only the fourth chart series shows its values with the 0.00E+00 scientific format via the Aspose.Cells API.
// Common Searches: Aspose.Cells how to set scientific notation for specific chart series data labels in C# | format fourth series data labels as 0.00E+00 in an Excel chart using Aspose.Cells | C# code to apply number format to chart series labels in an existing workbook | change data label display for series index 3 in Aspose.Cells chart | apply scientific number format to Excel chart series with Aspose.Cells .NET
// Tags: set chart series data label number format Aspose.Cells | apply scientific notation to Excel chart labels C# | fourth series data label formatting Aspose.Cells | chart series number format 0.00E+00 Aspose.Cells | enable data labels for specific series .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing workbook, accesses the first chart, verifies it contains at least four series, enables data labels on the fourth series, applies the scientific notation format "0.00E+00" to those labels, and saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the first chart on the worksheet (adjust index if needed)
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("Error: No charts found on the first worksheet.");
                return;
            }

            Chart chart = worksheet.Charts[0];

            // Verify that the chart has at least four series
            if (chart.NSeries.Count >= 4)
            {
                // The fourth series has zero‑based index 3
                Series fourthSeries = chart.NSeries[3];

                // Ensure data labels are displayed
                fourthSeries.DataLabels.ShowValue = true;

                // Apply scientific notation number format to the data labels
                // Format string "0.00E+00" displays numbers like 1.23E+04
                fourthSeries.DataLabels.NumberFormat = "0.00E+00";
            }
            else
            {
                Console.WriteLine("Warning: Chart does not contain at least four series. No changes applied.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
