// Title: Set custom font colors for chart legend entries by series index and keep legend background transparent using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change legend entry font color based on series index | How to make legend entry background transparent in an Aspose.Cells chart | Apply different colors to each legend item in a column chart with Aspose.Cells .NET | Customize chart legend text color while preserving transparent fill using Aspose.Cells API | Set legend entry Font.Color and BackgroundMode programmatically in Aspose.Cells
// Tags: legend entry Font.Color property Aspose.Cells | BackgroundMode.Transparent for legend entries | apply series-based colors to chart legend | Aspose.Cells chart legend text styling | column chart legend customization .NET

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace LegendCustomColorDemo
{
    // The example creates a workbook with three data series, adds a column chart, and iterates over each series to assign a distinct font color to its legend entry from a predefined palette while setting the legend entry's background mode to transparent, then saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // Populate sample data for three series
                // ------------------------------------------------------------
                // Categories
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");

                // Series 1
                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["B3"].PutValue(50);
                sheet.Cells["B4"].PutValue(70);

                // Series 2
                sheet.Cells["C1"].PutValue("Series 2");
                sheet.Cells["C2"].PutValue(40);
                sheet.Cells["C3"].PutValue(60);
                sheet.Cells["C4"].PutValue(80);

                // Series 3
                sheet.Cells["D1"].PutValue("Series 3");
                sheet.Cells["D2"].PutValue(20);
                sheet.Cells["D3"].PutValue(45);
                sheet.Cells["D4"].PutValue(65);

                // ------------------------------------------------------------
                // Add a column chart
                // ------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Add each series to the chart
                // Note: The first series (index 0) will be B2:B4, second C2:C4, third D2:D4
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries[0].Name = "Series 1";

                chart.NSeries.Add("C2:C4", true);
                chart.NSeries[1].Name = "Series 2";

                chart.NSeries.Add("D2:D4", true);
                chart.NSeries[2].Name = "Series 3";

                // Set category (X‑axis) data
                chart.NSeries.CategoryData = "A2:A4";

                // ------------------------------------------------------------
                // Apply custom font colors to legend entries based on series index
                // Keep the legend entry fill transparent
                // ------------------------------------------------------------
                Color[] palette = new Color[]
                {
                    Color.DarkBlue,
                    Color.DarkGreen,
                    Color.DarkRed,
                    Color.Purple,
                    Color.Orange
                };

                for (int i = 0; i < chart.NSeries.Count; i++)
                {
                    LegendEntry entry = chart.NSeries[i].LegendEntry;

                    // Set the font color – pick from the palette (wrap around if more series than colors)
                    entry.Font.Color = palette[i % palette.Length];

                    // Ensure the legend entry background is transparent
                    entry.BackgroundMode = BackgroundMode.Transparent;

                    // Optional: keep text fill enabled (default). If you ever set IsTextNoFill, set it to false.
                    entry.IsTextNoFill = false;
                }

                // ------------------------------------------------------------
                // Save the workbook
                // ------------------------------------------------------------
                string outFile = "LegendCustomColorDemo.xlsx";
                workbook.Save(outFile, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
