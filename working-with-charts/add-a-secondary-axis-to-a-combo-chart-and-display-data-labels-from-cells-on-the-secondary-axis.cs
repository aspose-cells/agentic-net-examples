// Title: Create a combo chart with a secondary axis and hide numeric values in data labels using Aspose.Cells for .NET (C#)
// Common Searches: how to add a secondary value axis to a combo chart with Aspose.Cells in C# | display data labels from worksheet cells on the secondary axis using Aspose.Cells | Aspose.Cells C# example for column and line series on primary and secondary axes | remove numeric values from data labels of a secondary series in an Excel chart with Aspose.Cells | save an Excel workbook that contains a combo chart with a secondary axis using Aspose.Cells
// Tags: secondary axis configuration for combo chart Aspose.Cells | line series on secondary axis Aspose.Cells | hide numeric values in chart data labels Aspose.Cells | populate chart series from worksheet cells Aspose.Cells | save workbook with combo chart Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, fills cells A1:C6 with month names and two data series, adds a combo chart (column for the primary series and line for the secondary series), assigns the secondary series to the secondary axis, disables numeric values in its data labels, and saves the file as ComboChartWithSecondaryAxis.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data
            ws.Cells["A1"].PutValue("Month");
            ws.Cells["B1"].PutValue("Primary");
            ws.Cells["C1"].PutValue("Secondary");

            ws.Cells["A2"].PutValue("Jan");
            ws.Cells["A3"].PutValue("Feb");
            ws.Cells["A4"].PutValue("Mar");
            ws.Cells["A5"].PutValue("Apr");
            ws.Cells["A6"].PutValue("May");

            ws.Cells["B2"].PutValue(10);
            ws.Cells["B3"].PutValue(20);
            ws.Cells["B4"].PutValue(30);
            ws.Cells["B5"].PutValue(40);
            ws.Cells["B6"].PutValue(50);

            ws.Cells["C2"].PutValue(100);
            ws.Cells["C3"].PutValue(200);
            ws.Cells["C4"].PutValue(300);
            ws.Cells["C5"].PutValue(400);
            ws.Cells["C6"].PutValue(500);

            // Add a combo chart (Column + Line)
            int chartIdx = ws.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = ws.Charts[chartIdx];

            // Primary series (Column) on primary axis
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries[0].Name = "Primary";

            // Secondary series (Line) – set series type to line
            chart.NSeries.Add("C2:C6", true);
            chart.NSeries[1].Name = "Secondary";
            chart.NSeries[1].Type = ChartType.Line;

            // Show data labels for the secondary series (hide numeric values)
            chart.NSeries[1].DataLabels.ShowValue = false;

            // Determine output file path and ensure directory exists
            string outputPath = "ComboChartWithSecondaryAxis.xlsx";
            string fullPath = Path.GetFullPath(outputPath);
            string outputDir = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
