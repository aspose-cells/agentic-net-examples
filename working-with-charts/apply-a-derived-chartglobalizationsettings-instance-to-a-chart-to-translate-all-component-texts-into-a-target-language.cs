// Title: Translating an Aspose.Cells .NET chart’s title, axis labels, and legend to Spanish using a custom dictionary
// Common Searches: C# code to localize Excel chart titles and axis labels with Aspose.Cells | how to change chart component text to another language in Aspose.Cells .NET | example of applying manual translations to Aspose.Cells chart elements | translate Excel chart legend and axis titles programmatically using Aspose.Cells | set Spanish labels for Aspose.Cells column chart in C#
// Tags: manual chart text translation Aspose.Cells | localize Excel chart components C# | set chart title language Aspose.Cells .NET | axis label localization Aspose.Cells | apply dictionary translations to chart objects

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example creates a workbook with month‑sales data, adds a column chart, defines an English‑to‑Spanish dictionary, and manually assigns the translated strings to the chart title, category axis title, value axis title, and legend before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Populate sample data
                var sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("January");
                sheet.Cells["A3"].PutValue("February");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["B3"].PutValue(1500);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                var chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B3", true);
                chart.NSeries.CategoryData = "A2:A3";

                // Define translations (English → Spanish)
                var translations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    { "Month", "Mes" },
                    { "Sales", "Ventas" },
                    { "January", "Enero" },
                    { "February", "Febrero" },
                    { "Chart Title", "Título del Gráfico" }
                };

                // Apply translations manually (custom globalization settings are not available in this version)
                if (translations.TryGetValue("Chart Title", out var chartTitle))
                {
                    chart.Title.Text = chartTitle;
                }

                // Translate axis titles if needed
                if (chart.CategoryAxis != null && translations.TryGetValue("Month", out var monthLabel))
                {
                    chart.CategoryAxis.Title.Text = monthLabel;
                }

                if (chart.ValueAxis != null && translations.TryGetValue("Sales", out var salesLabel))
                {
                    chart.ValueAxis.Title.Text = salesLabel;
                }

                // Determine output path and ensure the directory exists
                string outputPath = "ChartGlobalization.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
