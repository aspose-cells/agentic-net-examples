// Title: Set the second chart series to a custom theme color in an existing XLSX workbook using Aspose.Cells for .NET
// AI Prompts: Load an XLSX file with Aspose.Cells, locate the first chart on the first worksheet, change the foreground color of its second series to a specific RGB value, and save the workbook. | Using Aspose.Cells for .NET, retrieve the second series of a chart, assign a custom palette color (e.g., Accent1), and write the updated workbook to a new XLSX file.
// Common Searches: Aspose.Cells C# change color of second series in an existing chart | how to apply custom RGB color to a specific chart series using Aspose.Cells | update chart series theme color in a loaded workbook with Aspose.Cells for .NET | C# code to modify second series foreground color in Excel chart and save file | set custom palette color for chart series in Aspose.Cells workbook
// Tags: Aspose.Cells set chart series color C# | modify second series foreground Aspose.Cells | custom theme palette for Excel chart series | load workbook update chart Aspose.Cells | save workbook after chart formatting C#

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program loads "input.xlsx", accesses the first worksheet's first chart, changes the foreground color of the second series to a custom RGB value, and saves the modified workbook as "output.xlsx" using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Get the first chart on the worksheet
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            Chart chart = sheet.Charts[0];

            // Verify that the chart has at least two series
            if (chart.NSeries.Count > 1)
            {
                // Retrieve the second series (index 1)
                Series secondSeries = chart.NSeries[1];

                // Set its color (example: use Accent1 theme color)
                secondSeries.Area.ForegroundColor = Color.FromArgb(0, 112, 192); // custom color or theme accent
            }
            else
            {
                Console.WriteLine("The chart does not contain a second series.");
            }

            // Save the workbook back to XLSX format
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
