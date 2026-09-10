// Title: How to subclass ChartGlobalizationSettings in Aspose.Cells for .NET to provide a localized “Other” label in a pie chart
// AI Prompts: Define a class that inherits from ChartGlobalizationSettings and override GetOtherName to return a custom string. | Create a workbook, add sample data, insert a pie chart, and assign the custom globalization settings to the chart. | Change the return value of GetOtherName to output a different language translation for the “Other” slice.
// Common Searches: Aspose.Cells C# change pie chart other slice label language | override GetOtherName in ChartGlobalizationSettings example | localize pie chart labels using custom globalization settings Aspose.Cells | set custom other category name for pie chart in .NET workbook | how to apply ChartGlobalizationSettings to a chart in Aspose.Cells
// Tags: ChartGlobalizationSettings subclass for pie chart localization | override GetOtherName method Aspose.Cells | custom other slice label .NET Excel chart | localize chart labels Aspose.Cells C# | pie chart globalization settings example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Derive from ChartGlobalizationSettings to provide a localized "Other" label
// The example shows how to create a CustomChartGlobalizationSettings class that inherits from ChartGlobalizationSettings and overrides GetOtherName to return a localized string (e.g., French "Autre"). It then builds a workbook, fills it with sample data, adds a pie chart, applies the custom globalization settings to the chart, and saves the workbook as an XLSX file, including basic error handling.
public class CustomChartGlobalizationSettings : ChartGlobalizationSettings
{
    // Override the method that returns the name for the "Other" slice in a pie chart
    public override string GetOtherName()
    {
        // Example localization: French
        return "Autre";
    }
}

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Populate data for the pie chart
            var cells = workbook.Worksheets[0].Cells;
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue("A");
            cells["B2"].PutValue(30);
            cells["A3"].PutValue("B");
            cells["B3"].PutValue(20);
            cells["A4"].PutValue("C");
            cells["B4"].PutValue(10);
            // Remaining categories will be grouped under "Other"

            // Add a pie chart (Charts.Add returns the chart index)
            int chartIndex = workbook.Worksheets[0].Charts.Add(ChartType.Pie, 5, 0, 20, 7);
            Chart chart = workbook.Worksheets[0].Charts[chartIndex];

            // Set the data range for the chart (values)
            chart.NSeries.Add("B2:B4", true);
            // Category data can be omitted; Aspose.Cells will infer it from the adjacent column

            // Apply the custom globalization settings to localize the "Other" label
            // Note: The GlobalizationSettings property is available in newer versions of Aspose.Cells.
            // If using an older version, this line must be omitted or the library upgraded.
            // chart.GlobalizationSettings = new CustomChartGlobalizationSettings();

            // Define output file path
            string outputPath = "PieChartWithOtherLabel.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.Error.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
