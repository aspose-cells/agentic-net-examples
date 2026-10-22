// Title: How to add a linear trendline with its equation to the line series of a combo (column‑line) chart using Aspose.Cells for .NET (C#)
// AI Prompts: Create a combo chart with a column series for sales and a line series for profit, then add a linear trendline to the profit series and display its equation using Aspose.Cells in C#. | Use reflection to access the Trendlines collection of a chart series and enable the DisplayEquation property when the API does not expose it directly in Aspose.Cells for .NET.
// Common Searches: asp.net add linear trendline to line series in combo chart using Aspose.Cells | display trendline equation on Aspose.Cells chart C# | how to use reflection to add trendline in Aspose.Cells | combo chart column and line series with trendline Aspose.Cells example | Aspose.Cells add trendline to profit series in Excel workbook
// Tags: linear trendline implementation Aspose.Cells | show trendline formula Aspose.Cells | combo chart column line series Aspose.Cells | reflection based Trendlines access Aspose.Cells | save chart as XLSX Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsTrendlineExample
{
    // The sample creates a workbook, fills it with month, sales, and profit data, builds a combo chart (column for sales, line for profit), converts the profit series to a line type, and then uses reflection to add a linear trendline with its equation displayed. The workbook is saved as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the combo chart
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["C1"].PutValue("Profit");

                string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
                double[] sales = { 12000, 15000, 13000, 17000, 16000 };
                double[] profit = { 3000, 3500, 3200, 4000, 3800 };

                for (int i = 0; i < months.Length; i++)
                {
                    sheet.Cells[i + 1, 0].PutValue(months[i]);   // Column A
                    sheet.Cells[i + 1, 1].PutValue(sales[i]);   // Column B
                    sheet.Cells[i + 1, 2].PutValue(profit[i]);  // Column C
                }

                // Add a Combo chart (Column + Line) to the worksheet
                // Position: from row 7, column 0 to row 25, column 10
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Add the column series (Sales)
                chart.NSeries.Add("B2:B6", true);
                chart.NSeries[0].Name = "Sales";

                // Add the line series (Profit)
                chart.NSeries.Add("C2:C6", true);
                chart.NSeries[1].Name = "Profit";

                // Set the second series to be displayed as a line
                Series profitSeries = chart.NSeries[1];
                profitSeries.Type = ChartType.Line;

                // Attempt to add a linear trendline to the profit series using reflection.
                // This works for versions that expose the Trendlines property.
                try
                {
                    var trendlinesProp = profitSeries.GetType().GetProperty("Trendlines");
                    if (trendlinesProp != null)
                    {
                        var trendlinesObj = trendlinesProp.GetValue(profitSeries);
                        var addMethod = trendlinesObj?.GetType().GetMethod("Add");
                        var trendlineObj = addMethod?.Invoke(trendlinesObj, new object[] { TrendlineType.Linear });
                        var displayEqProp = trendlineObj?.GetType().GetProperty("DisplayEquation");
                        if (displayEqProp != null)
                        {
                            displayEqProp.SetValue(trendlineObj, true);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // If any reflection step fails, continue without a trendline.
                    Console.WriteLine("Trendline could not be added: " + ex.Message);
                }

                // Optional chart title
                chart.Title.Text = "Sales vs Profit with Trendline";

                // Determine output path and ensure the directory exists
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComboChartWithTrendline.xlsx");
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to: {outputPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine("Failed to save workbook: " + saveEx.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
