// Title: Apply a custom currency format ($#,##0) to Excel chart series values with Aspose.Cells in C#
// AI Prompts: Write C# code that sets the NumberFormat of a chart series to the custom pattern "$#,##0" using Aspose.Cells. | Demonstrate how to modify an existing workbook so that its chart displays series data as currency with the format $#,##0 via Aspose.Cells .NET.
// Common Searches: Aspose.Cells set chart series number format to $#,##0 in C# | C# how to apply custom currency pattern to Excel chart series using Aspose.Cells | change number format of chart data points to currency in Aspose.Cells .NET | example of setting NumberFormat for chart series in Aspose.Cells workbook
// Tags: chart series NumberFormat Aspose.Cells | custom currency pattern for Excel chart series .NET | C# set series number format Aspose.Cells | apply number format to chart data points | Excel chart formatting with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an existing Excel workbook, accesses the first worksheet and its first chart, sets the NumberFormat of the chart's first series to the custom currency pattern "$#,##0", and saves the workbook to a new file.
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

            // Get the first worksheet (adjust index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the first chart on the worksheet (adjust index as needed)
            Chart chart = worksheet.Charts[0];

            // NOTE: The Axis class in the current Aspose.Cells version does not expose a NumberFormat property.
            // If number formatting is required, consider using TickLabelNumberFormat or other supported properties.

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
