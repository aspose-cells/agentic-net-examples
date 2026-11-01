// Title: Set a custom palette color for the second series of the first chart in an existing XLSX workbook using Aspose.Cells for .NET
// AI Prompts: Open input.xlsx with Aspose.Cells, locate the first worksheet’s first chart, change the second series’ Area.ForegroundColor to a chosen Color (e.g., Red), and save the workbook as output.xlsx. | Using Aspise.Cells for .NET, verify that a chart contains at least two series, assign a custom palette color to the second series, and write the modified workbook back to an XLSX file.
// Common Searches: Aspose.Cells C# change color of second chart series in existing Excel file | set custom palette for a specific chart series using Aspose.Cells .NET | programmatically modify chart series foreground color in XLSX with Aspose | how to update chart series theme color without recreating chart Aspose.Cells
// Tags: Aspose.Cells chart series foreground color | second series color update Excel C# | apply custom series color Aspose.Cells | modify chart series theme Aspose.Cells | save workbook after chart color change .NET

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing XLSX file, accesses the first worksheet and its first chart, checks that the chart has at least two series, sets the foreground color of the second series to a custom color (e.g., Red), and saves the modified workbook as a new XLSX file, handling missing files and errors gracefully.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume the first worksheet contains the chart
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet has at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            // Get the first chart (index 0)
            Chart chart = sheet.Charts[0];

            // Ensure there are at least two series in the chart
            if (chart.NSeries.Count >= 2)
            {
                // Get the second series (index 1)
                Series secondSeries = chart.NSeries[1];

                // Set the series color to a custom color (e.g., Red)
                secondSeries.Area.ForegroundColor = Color.Red;
            }
            else
            {
                Console.WriteLine("The chart does not contain a second series.");
            }

            // Save the workbook back to XLSX
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
