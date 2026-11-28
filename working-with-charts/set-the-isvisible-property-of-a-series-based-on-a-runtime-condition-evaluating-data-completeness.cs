// Title: Hide or remove chart series with missing cells at runtime using Aspose.Cells for .NET (C#)
// Common Searches: how to hide a chart series that contains blank cells using Aspose.Cells in C# | remove incomplete series from an Excel column chart at runtime with Aspose.Cells | programmatically check for empty cells in a series range before displaying a chart in .NET | Aspose.Cells C# hide chart series based on data completeness condition | dynamic chart series visibility based on cell values using Aspose.Cells
// Tags: filter out incomplete chart series Aspose.Cells | blank cell detection in chart series C# | conditional series visibility Excel Aspose.Cells | runtime validation of chart data Aspose.Cells | chart series data completeness check Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SeriesVisibilityDemo
{
    // The example creates a workbook, adds sample data with one series containing blank cells, builds a column chart, then iterates through each series, parses its data range, checks every cell for emptiness, and removes any series that has incomplete data before saving the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data (some cells intentionally left blank to simulate incompleteness)
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Series 1");
                cells["C1"].PutValue("Series 2");

                cells["A2"].PutValue("Q1");
                cells["A3"].PutValue("Q2");
                cells["A4"].PutValue("Q3");
                cells["A5"].PutValue("Q4");

                // Series 1 – complete data
                cells["B2"].PutValue(10);
                cells["B3"].PutValue(20);
                cells["B4"].PutValue(30);
                cells["B5"].PutValue(40);

                // Series 2 – incomplete data (C4 is left blank)
                cells["C2"].PutValue(15);
                cells["C3"].PutValue(25);
                // cells["C4"] left blank on purpose
                cells["C5"].PutValue(45);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Add both series to the chart
                chart.NSeries.Add("B2:C5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Evaluate each series for data completeness and remove if incomplete
                for (int i = chart.NSeries.Count - 1; i >= 0; i--)
                {
                    Series series = chart.NSeries[i];
                    string valuesRange = series.Values; // e.g., "B2:B5" or "C2:C5"

                    if (string.IsNullOrEmpty(valuesRange))
                    {
                        // No data range – remove the series
                        chart.NSeries.RemoveAt(i);
                        Console.WriteLine($"Series \"{series.Name ?? $"Series{i + 1}"}\" removed (no data range).");
                        continue;
                    }

                    // Parse the range into a CellArea (handle single‑cell ranges as well)
                    string[] parts = valuesRange.Split(':');
                    CellArea area;
                    if (parts.Length == 2)
                    {
                        area = CellArea.CreateCellArea(parts[0], parts[1]);
                    }
                    else
                    {
                        area = CellArea.CreateCellArea(parts[0], parts[0]);
                    }

                    bool isComplete = true;

                    // Check each cell in the range for emptiness
                    for (int row = area.StartRow; row <= area.EndRow && isComplete; row++)
                    {
                        for (int col = area.StartColumn; col <= area.EndColumn && isComplete; col++)
                        {
                            Cell cell = cells[row, col];
                            if (cell == null || string.IsNullOrWhiteSpace(cell.StringValue))
                            {
                                isComplete = false;
                            }
                        }
                    }

                    if (!isComplete)
                    {
                        // Remove incomplete series
                        chart.NSeries.RemoveAt(i);
                        Console.WriteLine($"Series \"{series.Name ?? $"Series{i + 1}"}\" removed (incomplete data).");
                    }
                    else
                    {
                        Console.WriteLine($"Series \"{series.Name ?? $"Series{i + 1}"}\" kept (complete data).");
                    }
                }

                // Save the workbook
                string outputPath = "SeriesVisibilityDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
