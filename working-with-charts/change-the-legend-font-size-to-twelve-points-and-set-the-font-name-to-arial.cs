// Title: How to change a chart legend’s font size to 12 points and set the font to Arial using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set chart legend font size to 12 | Change Excel chart legend font to Arial with Aspose.Cells .NET | Programmatically modify legend font properties in an Excel workbook using Aspose.Cells | C# code to update chart legend styling in existing XLSX file | Aspose.Cells example for customizing chart legend appearance
// Tags: chart legend font customization Aspose.Cells | set legend font size .NET | apply Arial font to chart legend C# | modify Excel chart legend formatting Aspose | legend font properties Excel workbook C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an existing Excel workbook, accesses the first worksheet and its first chart, retrieves the chart's legend, sets the legend font size to 12 points and the font name to Arial, and then saves the workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Access the first chart on the worksheet
            Chart chart = sheet.Charts[0];

            // Retrieve the legend object of the chart
            Legend legend = chart.Legend;

            // Set the legend font size to 12 points
            legend.Font.Size = 12;

            // Set the legend font name to Arial
            legend.Font.Name = "Arial";

            // Save the workbook with the updated legend formatting
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
