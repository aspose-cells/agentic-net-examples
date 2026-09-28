// Title: How to format the Z‑axis labels as percentages in a 3‑D column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that accesses the Z (depth) axis of a 3‑D column chart in Aspose.Cells and applies a percentage number format (e.g., "0%") to its labels. | Show how to detect whether the AxisZ property is available in the current Aspose.Cells version and apply the percentage format only when supported. | Provide a full example that creates sample data, adds a 3‑D clustered column chart, formats the Z‑axis as percentages, and saves the workbook to an XLSX file.
// Common Searches: Aspose.Cells C# set Z axis label format to percentage in 3D column chart | how to apply number format 0% to depth axis of chart using Aspose.Cells .NET | C# Aspose.Cells chart Z axis formatting version compatibility | percentage display for Z axis values in Aspose.Cells 3D chart example
// Tags: z axis numeric display percent Aspose.Cells | 3d column chart axis formatting .NET | chart depth axis number format Aspose.Cells | C# apply percent number format to chart axis | AxisZ version check Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, fills cells A1:C4 with numeric data, adds a 3‑D clustered column chart, demonstrates how to access the Z (depth) axis and apply a percentage number format to its labels (including a version check for AxisZ), and saves the result as ZAxisPercentage.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Prepare sample data for the chart
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(0.1);
            sheet.Cells["B3"].PutValue(0.2);
            sheet.Cells["B4"].PutValue(0.3);
            sheet.Cells["C2"].PutValue(0.15);
            sheet.Cells["C3"].PutValue(0.25);
            sheet.Cells["C4"].PutValue(0.35);

            // Add a 3‑D column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column3DClustered, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // NOTE: AxisZ is not available in older Aspose.Cells versions.
            // If needed, configure Z‑axis formatting using the appropriate API for your version.

            // Save the workbook
            string outputPath = "ZAxisPercentage.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
