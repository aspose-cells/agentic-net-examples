// Title: Hide Zero‑Value Chart Series Across All Charts in a Workbook with Aspose.Cells for .NET
// Description: Demonstrates how to use Aspose.Cells' chart‑filtering API to scan every worksheet, evaluate each series' data range, and set Series.IsFiltered to hide series that contain only zero values before saving the workbook.
// Keywords: Aspose.Cells .NET | chart series filter | hide zero values | Series.IsFiltered | Excel chart automation | multiple charts workbook | chart data range parsing | remove empty series | column chart Aspose | line chart Aspose
// Common Searches: Aspose.Cells hide chart series with zeros | filter out zero‑value series in Excel using .NET | Series.IsFiltered example Aspose.Cells | iterate charts and hide empty series | chart data range check Aspose.Cells | remove blank series from multiple charts
// Developer Intent: Automatically conceal any chart series that consist solely of zero values in every chart of a workbook.
// Use Cases: Clean financial dashboards by removing series that have no data, ensuring charts display only relevant information. | Generate Excel reports where some categories may be empty, and automatically hide those empty series across different chart types. | Prepare presentation‑ready workbooks by filtering out zero‑value series from column, line, or other charts before distribution.
// AI Prompts: Write a reusable C# method that accepts a Workbook object and hides all chart series whose data range contains only zeros using Aspose.Cells. | Explain the behavior of Series.IsFiltered in Aspose.Cells and how to safely evaluate numeric cells while ignoring blanks or text. | Provide sample code that loops through all worksheets and charts, checks each series for all‑zero values, applies the filter, and saves the file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartFilteringDemo
{
    // Demonstrates how to use Aspose.Cells' chart‑filtering API to scan every worksheet, evaluate each series' data range, and set Series.IsFiltered to hide series that contain only zero values before saving the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create a new workbook ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ---------- Populate sample data ----------
                // Category column
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["A5"].PutValue("D");

                // Series 1 (contains zeros)
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(0);
                sheet.Cells["B3"].PutValue(0);
                sheet.Cells["B4"].PutValue(0);
                sheet.Cells["B5"].PutValue(0);

                // Series 2 (non‑zero values)
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(10);
                sheet.Cells["C3"].PutValue(20);
                sheet.Cells["C4"].PutValue(30);
                sheet.Cells["C5"].PutValue(40);

                // Series 3 (mixed values)
                sheet.Cells["D1"].PutValue("Series3");
                sheet.Cells["D2"].PutValue(5);
                sheet.Cells["D3"].PutValue(0);
                sheet.Cells["D4"].PutValue(15);
                sheet.Cells["D5"].PutValue(0);

                // ---------- Add two charts that use the same data ----------
                int chartIdx1 = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 8);
                Chart chart1 = sheet.Charts[chartIdx1];
                chart1.NSeries.Add("B2:B5", true);
                chart1.NSeries.Add("C2:C5", true);
                chart1.NSeries.Add("D2:D5", true);
                chart1.NSeries.CategoryData = "A2:A5";

                int chartIdx2 = sheet.Charts.Add(ChartType.Line, 7, 9, 20, 17);
                Chart chart2 = sheet.Charts[chartIdx2];
                chart2.NSeries.Add("B2:B5", true);
                chart2.NSeries.Add("C2:C5", true);
                chart2.NSeries.Add("D2:D5", true);
                chart2.NSeries.CategoryData = "A2:A5";

                // ---------- Filter out series that consist only of zero values ----------
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    foreach (Chart chart in ws.Charts)
                    {
                        // Iterate through each series in the chart
                        for (int i = 0; i < chart.NSeries.Count; i++)
                        {
                            Series series = chart.NSeries[i];

                            // Parse the series data range (e.g., "B2:B5")
                            CellArea area;
                            try
                            {
                                string[] parts = series.Values.Split(':');
                                if (parts.Length != 2)
                                    continue; // Invalid range format

                                area = CellArea.CreateCellArea(parts[0], parts[1]);
                            }
                            catch
                            {
                                // If the range cannot be parsed, skip this series
                                continue;
                            }

                            bool allZero = true;

                            // Examine each cell in the range
                            for (int row = area.StartRow; row <= area.EndRow && allZero; row++)
                            {
                                for (int col = area.StartColumn; col <= area.EndColumn && allZero; col++)
                                {
                                    Cell cell = ws.Cells[row, col];

                                    // Consider only numeric cells; non‑numeric cells break the zero condition
                                    if (cell.Type == CellValueType.IsNumeric)
                                    {
                                        if (cell.DoubleValue != 0)
                                            allZero = false;
                                    }
                                    else
                                    {
                                        // Non‑numeric (e.g., blank) also means the series is not "all zero"
                                        allZero = false;
                                    }
                                }
                            }

                            // If every data point is zero, hide the series
                            if (allZero)
                            {
                                series.IsFiltered = true;
                            }
                        }
                    }
                }

                // ---------- Save the workbook ----------
                string outputPath = "ChartSeriesFiltered.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
