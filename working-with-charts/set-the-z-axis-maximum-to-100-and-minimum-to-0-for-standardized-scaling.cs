// Title: Set manual Z‑axis minimum and maximum (0‑100) for a 3‑D clustered column chart using Aspose.Cells in C#
// AI Prompts: Create a 3‑D clustered column chart in an Excel workbook and fix its Z‑axis range to 0‑100 with Aspose.Cells for .NET. | Programmatically disable automatic scaling and assign specific min and max values to the series axis of a 3‑D chart using C#. | Generate an Excel file containing sample data and a 3‑D column chart whose Z‑axis limits are explicitly set to 0 and 100.
// Common Searches: Aspose.Cells C# set fixed Z axis range for 3D column chart | how to disable automatic scaling of series axis in Aspose.Cells | manual min max values for Excel 3D chart axis using .NET | C# example for setting Z‑axis limits in Aspose.Cells chart | standardize Z axis scaling in 3‑D clustered column chart Aspose
// Tags: Aspose.Cells series axis manual scaling | 3D column chart Z axis limits C# | set chart series axis min max Aspose.Cells | Excel 3D chart axis range .NET | disable automatic chart axis scaling Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a new workbook, adds sample data, inserts a 3‑D clustered column chart, and manually sets the Z‑axis (Series axis) minimum to 0 and maximum to 100 before saving the workbook as ChartWithZAxisScaling.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C2"].PutValue(40);
            sheet.Cells["C3"].PutValue(50);
            sheet.Cells["C4"].PutValue(60);

            // Add a 3‑D clustered column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column3DClustered, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:C4", true);

            // Access the Series (Z) axis and set manual scaling
            Axis seriesAxis = chart.SeriesAxis;
            if (seriesAxis != null)
            {
                seriesAxis.IsAutomaticMaxValue = false; // disable automatic max
                seriesAxis.IsAutomaticMinValue = false; // disable automatic min
                seriesAxis.MaxValue = 100;               // set maximum to 100
                seriesAxis.MinValue = 0;                 // set minimum to 0
            }

            // Ensure the output directory exists
            string outputPath = "ChartWithZAxisScaling.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
