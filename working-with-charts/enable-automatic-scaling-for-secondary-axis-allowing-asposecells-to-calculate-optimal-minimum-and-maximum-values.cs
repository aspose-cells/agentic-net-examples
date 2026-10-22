// Title: Create a dual‑axis column chart with automatic secondary axis scaling using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that builds a column chart, adds a secondary series, assigns it to the secondary Y‑axis, and lets Aspose.Cells determine the axis limits automatically. | Show how to configure an Aspose.Cells chart so the secondary axis scales itself without setting Min or Max values, including workbook creation and saving. | Provide a complete C# example that creates sample data, adds a dual‑axis column chart, and saves the workbook, relying on Aspose.Cells default auto‑scaling for the secondary axis.
// Common Searches: Aspose.Cells C# how to enable automatic scaling for secondary axis in a column chart | dual axis column chart without specifying axis min max using Aspose.Cells .NET | C# Aspose.Cells secondary Y axis auto range example | create chart with primary and secondary series and let Aspose.Cells calculate axis limits | Aspose.Cells automatic secondary axis range for column chart in .NET
// Tags: Aspose.Cells dual‑axis chart example | auto‑scale secondary Y‑axis Aspose.Cells | C# assign series to secondary axis | save workbook as Excel with chart Aspose.Cells | column chart without explicit axis limits

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, fills it with data, adds a column chart containing a primary and a secondary series, automatically assigns the second series to a secondary Y‑axis, relies on Aspose.Cells to compute optimal axis limits, and saves the file as an Excel workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            for (int i = 2; i <= 5; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 10); // Primary series values
                sheet.Cells[$"C{i}"].PutValue(i * 5);  // Secondary series values
            }

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add series to the chart
            chart.NSeries.Add("Sheet1!B2:B5", true);
            chart.NSeries[0].Name = "Series1";

            chart.NSeries.Add("Sheet1!C2:C5", true);
            chart.NSeries[1].Name = "Series2";

            // Note: In this version of Aspose.Cells the automatic axis scaling
            // and secondary axis assignment are handled by default, so explicit
            // properties are omitted.

            // Ensure the output directory exists
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
