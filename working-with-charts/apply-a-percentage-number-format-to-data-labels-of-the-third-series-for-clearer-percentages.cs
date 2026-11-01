// Title: Format the data labels of the third series as percentages in an Excel chart using Aspose.Cells for .NET (C#)
// AI Prompts: Use the Aspose.Cells C# API to set the NumberFormat of the third series' DataLabels to "0%" and enable ShowValue. | Load an existing workbook, locate the first chart, access series index 2, apply a percentage number format to its data labels, and save the file. | Programmatically change the data label display of a specific chart series to show percentages with Aspose.Cells and persist the changes.
// Common Searches: Aspose.Cells C# format third chart series data labels as percentage | set number format for a specific series data labels in Excel chart using Aspose.Cells | C# change data label display to 0% for a particular series in an Excel workbook | how to show percentage values on chart series data labels with Aspose.Cells .NET
// Tags: Aspose.Cells chart series data label percentage format | C# set data label number format Aspose.Cells | apply 0% format to third series chart Aspose.Cells | Excel chart data labels show value Aspose.Cells | modify existing workbook chart series Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example loads an existing Excel workbook, accesses the first worksheet's first chart, verifies that at least three series exist, then sets the third series' data labels to display values with a "0%" number format and enables label visibility before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart
            Chart chart = worksheet.Charts[0];

            // Ensure there are at least three series in the chart
            if (chart.NSeries.Count > 2)
            {
                // Access the third series (zero‑based index)
                Series thirdSeries = chart.NSeries[2];

                // Apply percentage number format to its data labels
                thirdSeries.DataLabels.NumberFormat = "0%";

                // Ensure data labels are displayed
                thirdSeries.DataLabels.ShowValue = true;
            }
            else
            {
                Console.WriteLine("The chart does not contain at least three series.");
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
