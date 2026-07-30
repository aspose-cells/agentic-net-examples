// Title: Aspose.Cells C# – Transparent Legend Fill for Negative Series in a Column Chart
// Description: This C# example creates an Excel workbook with a column chart that includes both positive and negative data. It scans each series, and when a series contains any negative values it sets the legend entry’s background to Transparent and optionally makes the legend text no‑fill using LegendEntry.BackgroundMode and IsTextNoFill. The workbook is then saved as an .xlsx file.
// Keywords: Aspose.Cells | C# | Excel chart | column chart | legend entry | transparent background | BackgroundMode.Transparent | IsTextNoFill | negative values | detect negative series | chart customization | automated reporting | Excel automation | Aspose.Cells API
// Common Searches: Aspose.Cells set legend entry transparent | C# make chart legend background transparent for negative values | How to hide legend fill for negative series in Aspose.Cells | LegendEntry.BackgroundMode Transparent example | Detect negative data in chart series Aspose.Cells
// Developer Intent: Apply a transparent fill to legend entries of chart series that contain any negative data points.
// Use Cases: Create financial dashboards where loss categories are shown without legend color to reduce visual clutter. | Generate compliance‑ready Excel reports that automatically hide legend colors for negative performance metrics. | Build automated spreadsheet generators that adapt legend styling based on data polarity. | Design presentation‑ready charts where negative series are distinguished by a clear legend entry.
// AI Prompts: Generate C# code with Aspose.Cells that iterates chart series, identifies negative values, and sets LegendEntry.BackgroundMode to Transparent and IsTextNoFill to true. | Show how to modify the legend entry border color for series with negative values using Aspose.Cells. | Explain how to extend the example to support line and bar charts in addition to column charts. | Provide a step‑by‑step guide to test the transparent legend feature in a unit test.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Alias to avoid conflict with System.Range (C# 8+)
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsLegendNegativeFill
{
    // This C# example creates an Excel workbook with a column chart that includes both positive and negative data. It scans each series, and when a series contains any negative values it sets the legend entry’s background to Transparent and optionally makes the legend text no‑fill using LegendEntry.BackgroundMode and IsTextNoFill. The workbook is then saved as an .xlsx file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (both positive and negative values)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");

                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(120);   // positive
                sheet.Cells["B3"].PutValue(-80);   // negative
                sheet.Cells["B4"].PutValue(150);   // positive

                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(-50);   // negative
                sheet.Cells["C3"].PutValue(70);    // positive
                sheet.Cells["C4"].PutValue(-30);   // negative

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Add the first series (Series1)
                chart.NSeries.Add("B2:B4", true);
                // Add the second series (Series2)
                chart.NSeries.Add("C2:C4", true);
                // Set category (X‑axis) data
                chart.NSeries.CategoryData = "A2:A4";

                // Iterate through each series to determine if it contains negative values
                for (int i = 0; i < chart.NSeries.Count; i++)
                {
                    Series series = chart.NSeries[i];
                    bool hasNegative = false;

                    // Get the data range for the series (e.g., "B2:B4")
                    string valuesRange = series.Values;
                    if (!string.IsNullOrEmpty(valuesRange))
                    {
                        AsposeRange range = sheet.Cells.CreateRange(valuesRange);
                        foreach (Cell cell in range)
                        {
                            if (cell.Value != null &&
                                double.TryParse(cell.Value.ToString(), out double d) &&
                                d < 0)
                            {
                                hasNegative = true;
                                break;
                            }
                        }
                    }

                    // If the series has any negative value, make its legend entry background transparent
                    if (hasNegative)
                    {
                        LegendEntry legendEntry = series.LegendEntry;
                        legendEntry.BackgroundMode = BackgroundMode.Transparent;
                        // Optionally, make the legend text transparent as well
                        legendEntry.IsTextNoFill = true;
                    }
                }

                // Define output file path
                string outputPath = "LegendNegativeFillDemo.xlsx";

                // Ensure the directory exists before saving
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
}
