// Title: Programmatically read a chart legend’s position and move it to the opposite corner with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# read current legend position of an Excel chart | how to switch Excel chart legend from top to bottom using Aspose.Cells | C# code to move chart legend to opposite corner in a workbook | Aspose.Cells change legend position enum to opposite value | example of toggling chart legend placement with Aspose.Cells .NET
// Tags: Aspose.Cells chart legend position manipulation | C# Aspose.Cells legend enum handling | Excel chart legend opposite corner switch | Aspose.Cells modify legend placement .NET | programmatic legend relocation in Excel workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace LegendPositionSwitcher
{
    // The sample loads an existing Excel file (or creates a new workbook with a column chart if the file is missing), accesses the first chart’s legend, reads its current LegendPositionType, determines the opposite corner via a helper method, updates the legend position, and saves the workbook as output.xlsx.
    class Program
    {
        // Returns the opposite legend position for the supported enum values.
        static LegendPositionType GetOppositePosition(LegendPositionType current)
        {
            switch (current)
            {
                case LegendPositionType.Top:
                    return LegendPositionType.Bottom;
                case LegendPositionType.Bottom:
                    return LegendPositionType.Top;
                case LegendPositionType.Left:
                    return LegendPositionType.Right;
                case LegendPositionType.Right:
                    return LegendPositionType.Left;
                // For positions not supported in this version (e.g., Custom, None), keep unchanged.
                default:
                    return current;
            }
        }

        static void Main(string[] args)
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            Workbook workbook = null;

            // Load existing workbook if it exists; otherwise create a sample workbook with a chart.
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load workbook \"{inputPath}\": {ex.Message}");
                    return;
                }
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a sample workbook.");

                try
                {
                    workbook = new Workbook();
                    Worksheet sheet = workbook.Worksheets[0];

                    // Populate sample data
                    sheet.Cells["A1"].PutValue("Category");
                    sheet.Cells["B1"].PutValue("Value");
                    sheet.Cells["A2"].PutValue("A");
                    sheet.Cells["B2"].PutValue(10);
                    sheet.Cells["A3"].PutValue("B");
                    sheet.Cells["B3"].PutValue(20);
                    sheet.Cells["A4"].PutValue("C");
                    sheet.Cells["B4"].PutValue(30);

                    // Add a chart
                    int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                    Chart chart = sheet.Charts[chartIndex];
                    chart.NSeries.Add("B2:B4", true);
                    chart.NSeries.CategoryData = "A2:A4";

                    // Ensure the legend is visible and set an initial position
                    chart.Legend.IsDeleted = false;
                    chart.Legend.Position = LegendPositionType.Right;

                    Console.WriteLine("Sample workbook created with a chart and legend.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to create sample workbook: {ex.Message}");
                    return;
                }
            }

            // Work with the first worksheet and its first chart
            Worksheet ws = workbook.Worksheets[0];
            if (ws.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet. Cannot modify legend position.");
                return;
            }

            Chart targetChart = ws.Charts[0];
            Legend legend = targetChart.Legend;

            // Read current legend position
            LegendPositionType currentPos = legend.Position;
            Console.WriteLine($"Current legend position: {currentPos}");

            // Determine opposite position
            LegendPositionType oppositePos = GetOppositePosition(currentPos);
            Console.WriteLine($"Switching legend position to: {oppositePos}");

            // Apply new position
            legend.Position = oppositePos;

            // Save the workbook
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved with updated legend position to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
