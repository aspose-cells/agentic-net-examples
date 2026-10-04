// Title: How to apply a custom currency number format to the data labels of the first chart series using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells to set series.DataLabels.NumberFormat to "$#,##0.00" for the first chart series in a workbook. | Programmatically format chart data labels as currency in C# with Aspose.Cells by modifying the NumberFormat property of the series' DataLabels. | Apply a custom currency pattern to all data labels of the first series in an Excel chart using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# set currency format for chart series data labels | How to change number format of data labels in the first series of an Excel chart using Aspose.Cells | Apply custom number format to chart data labels in a .NET workbook
// Tags: Aspose.Cells chart series data label number format | C# currency number format for Excel chart data labels | apply custom number format to chart series Aspose.Cells | first series data label formatting .NET | Excel chart data label currency pattern Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Loads an Excel workbook, accesses the first worksheet's first chart, sets the NumberFormat of the first series' data labels to the currency pattern "$#,##0.00", and saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook that contains a chart
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count > 0)
            {
                // Access the first chart
                Chart chart = sheet.Charts[0];

                // Ensure the chart has at least one series
                if (chart.NSeries.Count > 0)
                {
                    // Get the first series
                    Series series = chart.NSeries[0];

                    // The following properties are available in newer Aspose.Cells versions.
                    // If they are not present in the referenced version, they are omitted.
                    // series.IsShowValue = true;
                    // series.DataLabels.IsValueShown = true;

                    // Apply a custom currency number format to the data labels of this series
                    series.DataLabels.NumberFormat = "$#,##0.00";
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
