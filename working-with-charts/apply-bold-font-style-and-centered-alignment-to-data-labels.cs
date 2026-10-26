// Title: Apply bold font style to chart data labels in an Aspose.Cells column chart using C#
// Common Searches: how to set chart data label font to bold with Aspose.Cells C# | Aspose.Cells C# make data labels bold in a column chart | C# Aspose.Cells enable and style data labels for chart series | format chart data labels as bold using Aspose.Cells .NET | display and customize data labels in an Aspose.Cells column chart C#
// Tags: Aspose.Cells chart data label bold styling | column chart data label formatting Aspose.Cells | C# chart data label appearance customization | Excel workbook chart label styling .NET | Aspose.Cells series label font configuration

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing workbook, adds a column chart, creates a series from range A1:A5, enables data labels, sets the data label font to bold, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Add a column chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Add a series using the values range (A1:A5)
            chart.NSeries.Add("A1:A5", false);

            // Show data labels (values) for the series
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Apply bold font style to the data labels
            chart.NSeries[0].DataLabels.Font.IsBold = true;

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
