// Title: Add a fallback to default globalization and axis label formatting when custom chart label logic throws an exception in Aspose.Cells for .NET
// AI Prompts: Wrap the SetCustomAxisLabels call in a try‑catch, and on catch reset workbook.Settings.CultureInfo to CultureInfo.InvariantCulture then invoke SetDefaultAxisLabels to apply a generic number format before saving. | Implement error handling that logs any exception from custom axis label generation, reverts the workbook's culture to invariant, and ensures the chart's CategoryAxis.TickLabels use the "General" format.
// Common Searches: Aspose.Cells how to revert to invariant culture after custom axis label exception | C# fallback globalization settings for chart axis labels in Aspose.Cells | Handling errors in custom chart label methods with Aspose.Cells .NET | Set default tick label number format when custom label method fails Aspose.Cells | Try‑catch around SetCustomAxisLabels to restore default culture in workbook
// Tags: fallback globalization Aspose.Cells .NET | reset workbook CultureInfo on chart label error | default axis tick label format Aspose.Cells | exception handling for custom chart labels | apply invariant culture to workbook Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Globalization;

// The example creates a workbook, applies French culture, adds data and a column chart, then attempts to set custom axis labels. If the custom method throws, the code logs the error, resets the workbook's CultureInfo to InvariantCulture, applies a default "General" number format to the chart's category axis, and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Apply custom globalization settings
            var customCulture = new CultureInfo("fr-FR");
            workbook.Settings.CultureInfo = customCulture;

            // Populate sample data
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("A1:A3", true);

            // Attempt to set custom axis labels; fallback on failure
            try
            {
                SetCustomAxisLabels(chart);
            }
            catch (Exception ex)
            {
                // Log the exception (optional)
                Console.WriteLine("Custom label method failed: " + ex.Message);

                // Revert to default globalization settings
                workbook.Settings.CultureInfo = CultureInfo.InvariantCulture;

                // Apply default axis label settings
                SetDefaultAxisLabels(chart);
            }

            // Save the workbook
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception e)
        {
            // General exception handling to prevent crashes
            Console.WriteLine("An error occurred: " + e.Message);
        }
    }

    // Custom label method that may throw an exception
    static void SetCustomAxisLabels(Chart chart)
    {
        // Example condition that triggers an exception
        if (chart.Worksheet.Workbook.Settings.CultureInfo.Name == "fr-FR")
        {
            throw new InvalidOperationException("Custom label generation failed for French culture.");
        }

        // If no exception, apply custom number format to axis labels
        chart.CategoryAxis.TickLabels.NumberFormat = "0.00";
    }

    // Default label method used as a fallback
    static void SetDefaultAxisLabels(Chart chart)
    {
        chart.CategoryAxis.TickLabels.NumberFormat = "General";
    }
}
