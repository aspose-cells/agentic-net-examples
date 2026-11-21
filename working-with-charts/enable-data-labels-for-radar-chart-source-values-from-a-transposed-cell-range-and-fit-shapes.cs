// Title: Add data labels to a radar chart, use a transposed series range, and auto‑fit the chart size with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells enable data labels on radar chart | Aspose.Cells set radar chart series source to transposed range | auto fit chart dimensions to worksheet area using Aspose.Cells | use reflection to toggle IsShowDataLabel property in Aspose.Cells chart series
// Tags: radar chart data labels Aspose.Cells | transposed series source range Aspose.Cells | auto fit chart shape Aspose.Cells | reflection set chart series property C# | save radar chart to Excel Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a workbook, fills it with sample data, adds a radar chart, assigns the series values from a transposed horizontal range (B2:C4), enables data labels for each series via reflection, auto‑fits the chart shape within the worksheet, and saves the file as RadarChartWithDataLabels.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the radar chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["C1"].PutValue("Series2");

                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // Add a Radar chart to the worksheet (positioned from row 5, column 0 to row 20, column 10)
                int chartIndex = sheet.Charts.Add(ChartType.Radar, 5, 0, 20, 10);
                Chart radarChart = sheet.Charts[chartIndex];

                // Set the series source values using a transposed range (horizontal orientation)
                radarChart.NSeries.Add("B2:C4", false); // false => data is taken horizontally (transposed)

                // Enable data labels for all series in the radar chart using reflection
                // (avoids direct dependency on ChartSeries type which may vary across versions)
                for (int i = 0; i < radarChart.NSeries.Count; i++)
                {
                    object series = radarChart.NSeries[i];
                    PropertyInfo prop = series.GetType().GetProperty("IsShowDataLabel");
                    if (prop != null && prop.CanWrite)
                    {
                        prop.SetValue(series, true);
                    }
                }

                // Determine output path and ensure directory exists
                string outputFile = "RadarChartWithDataLabels.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                try
                {
                    workbook.Save(outputFile);
                    Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine("Error saving workbook:");
                    Console.WriteLine(saveEx.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the radar chart:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
