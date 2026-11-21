// Title: How to set the primary Y‑axis major unit to 10 in an Aspose.Cells column chart using C#
// Common Searches: Aspose.Cells C# set primary value axis major unit to 10 | C# code to change Y axis major grid interval in Aspose.Cells chart | adjust chart Y‑axis major unit for column chart with Aspose.Cells | standardize grid lines across series in Aspose.Cells column chart | set chart value axis major unit property in C# Aspose.Cells
// Tags: primary value axis major unit Aspose.Cells | chart Y axis interval C# | Aspose.Cells column chart axis scaling | configure chart grid interval Aspose.Cells | set major unit property Aspose.Cells chart

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, adds a column chart, sets the primary Y (value) axis major unit to 10 to align grid lines across all series, and saves the file as ChartWithMajorUnit.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a column chart (example placement)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the major unit interval of the primary Y (value) axis to 10
            chart.ValueAxis.MajorUnit = 10;

            // Save the workbook
            string outputPath = "ChartWithMajorUnit.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
