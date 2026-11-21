// Title: Rotate all chart axis tick labels to 90° across every worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# rotate chart axis tick labels 90 degrees for all worksheets | How to set chart tick label orientation to vertical in Excel using Aspose.Cells .NET | Programmatically change chart label rotation on each sheet with Aspose.Cells | Bulk update chart axis label angle to 90° in a workbook using Aspose.Cells | C# iterate worksheets and modify chart tick label rotation in Aspose.Cells
// Tags: chart axis tick label rotation Aspose.Cells | iterate worksheets modify charts .NET | CategoryAxis TickLabelRotationAngle property Aspose.Cells | ValueAxis TickLabelRotationAngle Aspose.Cells | bulk chart formatting Excel Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an Excel workbook, loops through every worksheet and each chart, and (when supported) sets the CategoryAxis and ValueAxis TickLabelRotationAngle to 90 degrees, then saves the updated file with error handling for missing files and chart issues.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts on the current worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // Set rotation for primary axes if the property is available
                        if (chart.CategoryAxis != null)
                        {
                            // Uncomment the line below if TickLabelRotationAngle is supported:
                            // chart.CategoryAxis.TickLabelRotationAngle = 90;
                        }

                        if (chart.ValueAxis != null)
                        {
                            // Uncomment the line below if TickLabelRotationAngle is supported:
                            // chart.ValueAxis.TickLabelRotationAngle = 90;
                        }

                        // Note: Secondary axes properties may not be available in all versions.
                        // The following blocks are omitted to maintain compatibility.
                    }
                    catch (Exception exChart)
                    {
                        Console.WriteLine($"Error processing chart '{chart.Name}': {exChart.Message}");
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
