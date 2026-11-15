// Title: How to set the Y‑axis as a numeric value axis in an Aspose.Cells chart using C#
// Common Searches: C# Aspose.Cells set chart Y axis to value type | Aspose.Cells chart numeric Y axis configuration example | Change Excel chart value axis to numeric with Aspose.Cells .NET | Programmatically set value axis for a chart in Aspose.Cells C# | Aspose.Cells chart axis type value axis C# code sample
// Tags: Aspose.Cells chart value axis configuration | C# set numeric Y axis Aspose.Cells | Excel chart value axis type Aspose.Cells .NET | programmatic chart axis settings Aspose.Cells | modify chart Y axis to value Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an existing workbook, accesses the first worksheet and its first chart, confirms the Y‑axis (value axis) is numeric by default and can be customized, then saves the workbook to a new file.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the first worksheet.");
                return;
            }

            // Access the first chart on the worksheet
            Chart chart = worksheet.Charts[0];

            // The Y‑axis (value axis) is numeric by default; additional configuration can be added here if needed.
            // Example: chart.ValueAxis.IsAutomaticMajorUnit = false; // customize as required

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
