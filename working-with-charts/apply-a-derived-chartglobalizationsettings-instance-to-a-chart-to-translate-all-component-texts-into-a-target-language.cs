// Title: Manually translate a column chart’s title, axis labels, and worksheet headers to French using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that builds a column chart from sample data and substitutes the chart title, category axis labels, and worksheet header cells with French equivalents from a provided dictionary. | Show how to iterate over worksheet cells and chart components to apply localized strings before saving the workbook as an .xlsx file. | Provide a snippet that creates the output directory if missing and saves the localized workbook, handling any exceptions.
// Common Searches: asp.net aspose.cells translate chart title to French | how to localize Excel chart axis labels with Aspose.Cells C# | example of replacing worksheet header text for multilingual Excel using Aspose.Cells | C# Aspose.Cells chart globalization without ChartGlobalizationSettings | save Excel workbook with localized chart components in .NET
// Tags: Aspose.Cells chart text localization C# | dictionary driven chart translation Aspose.Cells | column chart multilingual labels .NET | Excel worksheet cell value replacement Aspose.Cells | output directory creation before workbook save C# | exception handling for Aspose.Cells workbook generation

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartGlobalization
{
    // The example creates a new workbook, adds month and sales data, inserts a column chart, defines an English‑to‑French translation dictionary, manually replaces the chart title, worksheet header cells, and month names on the category axis with their French equivalents, ensures the output directory exists, and saves the workbook as ChartWithGlobalization.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                var workbook = new Workbook();

                // Add sample data to the first worksheet.
                var sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("January");
                sheet.Cells["A3"].PutValue("February");
                sheet.Cells["A4"].PutValue("March");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["B3"].PutValue(1500);
                sheet.Cells["B4"].PutValue(1800);

                // Insert a column chart.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                var chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";
                chart.Title.Text = "Monthly Sales";

                // Prepare translations (e.g., English → French).
                var translations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    { "Monthly Sales", "Ventes Mensuelles" },
                    { "Month", "Mois" },
                    { "Sales", "Ventes" },
                    { "January", "Janvier" },
                    { "February", "Février" },
                    { "March", "Mars" }
                };

                // Apply translations manually.
                // Translate chart title.
                if (translations.TryGetValue(chart.Title.Text, out var titleTrans))
                {
                    chart.Title.Text = titleTrans;
                }

                // Translate header cells.
                foreach (var cellRef in new[] { "A1", "B1" })
                {
                    var cell = sheet.Cells[cellRef];
                    if (cell.Value != null && translations.TryGetValue(cell.StringValue, out var headerTrans))
                    {
                        cell.PutValue(headerTrans);
                    }
                }

                // Translate month names in the category axis.
                for (int row = 2; row <= 4; row++)
                {
                    var cell = sheet.Cells[$"A{row}"];
                    if (cell.Value != null && translations.TryGetValue(cell.StringValue, out var monthTrans))
                    {
                        cell.PutValue(monthTrans);
                    }
                }

                // Ensure output directory exists.
                string outputPath = "ChartWithGlobalization.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
