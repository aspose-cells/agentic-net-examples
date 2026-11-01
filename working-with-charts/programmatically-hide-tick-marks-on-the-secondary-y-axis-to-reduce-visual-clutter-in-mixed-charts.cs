// Title: Hide tick marks and axis line on the secondary Y‑axis of a mixed column‑line chart with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that builds a mixed column and line chart using Aspose.Cells, moves the line series to the secondary Y‑axis, and disables major/minor tick marks and the axis line on that secondary axis. | Demonstrate how to employ reflection to set IsSecondaryAxis (or PlotOnSecondAxis) and hide secondary axis tick marks for compatibility across different Aspose.Cells versions.
// Common Searches: Aspose.Cells C# hide secondary Y axis tick marks in mixed column line chart | remove major and minor tick marks from secondary value axis Aspose.Cells .NET | programmatically assign line series to secondary axis Aspose.Cells | reflection to set IsSecondaryAxis property Aspose.Cells version compatibility | how to hide axis line on secondary value axis in Aspose.Cells chart
// Tags: hide secondary axis tick marks Aspose.Cells | mixed column and line chart Aspose.Cells C# | assign series to secondary Y axis Aspose.Cells | reflection for axis property compatibility Aspose.Cells | remove secondary axis line Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, populates sample data, adds a mixed column‑line chart, assigns the line series to the secondary Y‑axis, and uses reflection to turn off major/minor tick marks and the axis line on that secondary axis before saving the file as an XLSX document.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet's cells
            var workbook = new Workbook();
            var cells = workbook.Worksheets[0].Cells;

            // Populate sample data
            cells["A1"].PutValue("Month");
            cells["B1"].PutValue("Primary");
            cells["C1"].PutValue("Secondary");
            cells["A2"].PutValue("Jan");
            cells["A3"].PutValue("Feb");
            cells["A4"].PutValue("Mar");
            cells["B2"].PutValue(10);
            cells["B3"].PutValue(20);
            cells["B4"].PutValue(30);
            cells["C2"].PutValue(100);
            cells["C3"].PutValue(200);
            cells["C4"].PutValue(300);

            // Add a mixed chart (column + line) to the first worksheet
            var sheet = workbook.Worksheets[0];
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            var chart = sheet.Charts[chartIndex];

            // Primary series (column) – first data column
            int primarySeriesIndex = chart.NSeries.Add("B2:B4", true);
            var primarySeries = chart.NSeries[primarySeriesIndex];
            primarySeries.Name = "Primary";

            // Secondary series (line) – second data column
            int secondarySeriesIndex = chart.NSeries.Add("C2:C4", true);
            var secondarySeries = chart.NSeries[secondarySeriesIndex];
            secondarySeries.Name = "Secondary";
            secondarySeries.Type = ChartType.Line;

            // Assign series to secondary axis using reflection (covers different API versions)
            var isSecProp = secondarySeries.GetType().GetProperty("IsSecondaryAxis");
            if (isSecProp != null && isSecProp.CanWrite)
            {
                isSecProp.SetValue(secondarySeries, true);
            }
            else
            {
                var plotOnSecondProp = secondarySeries.GetType().GetProperty("PlotOnSecondAxis");
                if (plotOnSecondProp != null && plotOnSecondProp.CanWrite)
                {
                    plotOnSecondProp.SetValue(secondarySeries, true);
                }
            }

            // Hide tick marks on the secondary Y axis (if supported) using reflection
            var secondaryAxisProp = chart.GetType().GetProperty("SecondaryValueAxis");
            if (secondaryAxisProp != null)
            {
                var secondaryAxis = secondaryAxisProp.GetValue(chart);
                var axisType = secondaryAxis.GetType();

                var majorTickProp = axisType.GetProperty("HasMajorTickMark");
                var minorTickProp = axisType.GetProperty("HasMinorTickMark");
                var axisLineProp = axisType.GetProperty("HasAxisLine");

                if (majorTickProp != null && majorTickProp.CanWrite) majorTickProp.SetValue(secondaryAxis, false);
                if (minorTickProp != null && minorTickProp.CanWrite) minorTickProp.SetValue(secondaryAxis, false);
                if (axisLineProp != null && axisLineProp.CanWrite) axisLineProp.SetValue(secondaryAxis, false);
            }

            // Define output file path
            string outputPath = "MixedChartWithHiddenSecondaryTicks.xlsx";

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
