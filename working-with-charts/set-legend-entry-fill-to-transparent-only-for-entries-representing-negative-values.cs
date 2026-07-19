// Title: Make chart legend entries transparent for series with negative values using Aspose.Cells for .NET
// Description: C# example that creates an Excel workbook, adds a column chart with mixed positive and negative data, scans each series for negative values, and applies a transparent background (BackgroundMode.Transparent) plus no‑fill text to the legend entries of those series before saving the file.
// Keywords: Aspose.Cells | C# chart legend transparent | BackgroundMode.Transparent | negative values legend entry | Excel column chart Aspose.Cells | .NET chart customization | legend entry fill | detect negative series | Aspose.Cells example GitHub
// Common Searches: Aspose.Cells set legend entry transparent | make legend background transparent for negative series C# | how to hide legend fill for specific chart series Aspose.Cells | transparent legend entry Excel chart .NET | detect negative values in chart series Aspose.Cells
// Developer Intent: Apply a transparent background to legend entries of chart series that contain any negative data points.
// Use Cases: Financial reports where negative performance series should be visually de‑emphasized in the legend. | Automated Excel generation that adapts legend styling based on data polarity. | Custom Excel dashboards that require conditional legend formatting without manual editing.
// AI Prompts: Generate C# code with Aspose.Cells that sets LegendEntry.BackgroundMode to Transparent only for series that have at least one negative value. | Explain step‑by‑step how to iterate chart series, check cell values for negativity, and apply LegendEntry.IsTextNoFill for a fully transparent legend entry. | Provide a complete Aspose.Cells example that creates a column chart, detects negative series, makes their legend entries transparent, and saves the workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsLegendNegativeFillDemo
{
    // C# example that creates an Excel workbook, adds a column chart with mixed positive and negative data, scans each series for negative values, and applies a transparent background (BackgroundMode.Transparent) plus no‑fill text to the legend entries of those series before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data with both positive and negative values
                // Series 1 (positive only)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(50);
                sheet.Cells["B3"].PutValue(70);
                sheet.Cells["B4"].PutValue(90);

                // Series 2 (contains negative values)
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(-30);
                sheet.Cells["C3"].PutValue(20);
                sheet.Cells["C4"].PutValue(-10);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Add both series to the chart
                chart.NSeries.Add("B2:B4", true); // Series 1
                chart.NSeries.Add("C2:C4", true); // Series 2
                chart.NSeries.CategoryData = "A2:A4";

                // Iterate through each series to determine if it contains negative values
                for (int i = 0; i < chart.NSeries.Count; i++)
                {
                    Series series = chart.NSeries[i];
                    bool hasNegative = false;

                    // Get the range that holds the series values (e.g., "B2:B4")
                    string valuesRange = series.Values;

                    // Ensure the range string is not empty
                    if (!string.IsNullOrEmpty(valuesRange))
                    {
                        // Iterate through each cell in the range
                        foreach (Cell cell in sheet.Cells.CreateRange(valuesRange))
                        {
                            if (cell.Value != null && double.TryParse(cell.Value.ToString(), out double val))
                            {
                                if (val < 0)
                                {
                                    hasNegative = true;
                                    break;
                                }
                            }
                        }
                    }

                    // If the series has any negative value, set its legend entry background to transparent
                    if (hasNegative)
                    {
                        LegendEntry legendEntry = series.LegendEntry;
                        legendEntry.BackgroundMode = BackgroundMode.Transparent;
                        // Make the legend text have no fill for full transparency effect
                        legendEntry.IsTextNoFill = true;
                    }
                }

                // Save the workbook
                string outputPath = "LegendNegativeFillDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
