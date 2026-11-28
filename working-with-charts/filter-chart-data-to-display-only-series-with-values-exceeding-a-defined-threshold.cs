// Title: C# example: Filter out column chart series in Aspose.Cells when all values are below a threshold
// Common Searches: Aspose.Cells C# hide column chart series with values below 25 | filter chart series by value threshold using Aspose.Cells | set IsFiltered on Excel chart series programmatically in C# | remove Excel chart series that do not meet minimum value Aspose.Cells | example code for chart series filtering based on data range Aspose.Cells
// Tags: Aspose.Cells chart series threshold filtering | C# hide Excel chart series Aspose.Cells | IsFiltered property Aspose.Cells chart | column chart data range analysis C# | chart series visibility control Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartSeriesFilterDemo
{
    // The sample creates a workbook, populates it with category and three data series, and adds a column chart. It then iterates each series, extracts its cell range, scans for any numeric value exceeding a defined threshold (25.0), and sets the series' IsFiltered flag to hide series that never exceed the threshold. The workbook is saved as an XLSX file and the chart is optionally exported to PDF.
    class Program
    {
        static void Main(string[] args)
        {
            const double Threshold = 25.0;

            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["C1"].PutValue("Series 2");
                sheet.Cells["D1"].PutValue("Series 3");

                string[] categories = { "Jan", "Feb", "Mar", "Apr", "May" };
                double[] series1 = { 10, 20, 30, 40, 50 };
                double[] series2 = { 5, 15, 25, 35, 45 };
                double[] series3 = { 2, 4, 6, 8, 10 };

                for (int i = 0; i < categories.Length; i++)
                {
                    int row = i + 2; // Excel rows are 1‑based; row 2 is first data row
                    sheet.Cells[$"A{row}"].PutValue(categories[i]);
                    sheet.Cells[$"B{row}"].PutValue(series1[i]);
                    sheet.Cells[$"C{row}"].PutValue(series2[i]);
                    sheet.Cells[$"D{row}"].PutValue(series3[i]);
                }

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Add three series to the chart
                chart.NSeries.Add("B2:B6", true); // Series 1
                chart.NSeries.Add("C2:C6", true); // Series 2
                chart.NSeries.Add("D2:D6", true); // Series 3
                chart.NSeries.CategoryData = "A2:A6";

                // Title for clarity
                chart.Title.Text = $"Series values > {Threshold}";

                // Iterate over each series and decide whether to filter it
                for (int i = 0; i < chart.NSeries.Count; i++)
                {
                    Series series = chart.NSeries[i];
                    string valuesFormula = series.Values; // e.g. "=Sheet1!$B$2:$B$6"

                    // Remove leading '=' and any sheet name, then strip '$'
                    string rangeAddress = valuesFormula.TrimStart('=');
                    if (rangeAddress.Contains("!"))
                        rangeAddress = rangeAddress.Split('!')[1];
                    rangeAddress = rangeAddress.Replace("$", "");

                    // Create a Range object from the address (use fully qualified type to avoid ambiguity)
                    Aspose.Cells.Range dataRange;
                    try
                    {
                        dataRange = workbook.Worksheets.CreateRange(rangeAddress, sheet.Index);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to create range '{rangeAddress}': {ex.Message}");
                        continue;
                    }

                    // Determine the bounds of the range
                    int startRow = dataRange.FirstRow;
                    int endRow = startRow + dataRange.RowCount - 1;
                    int startCol = dataRange.FirstColumn;
                    int endCol = startCol + dataRange.ColumnCount - 1;

                    bool keepSeries = false;

                    // Scan all cells in the range
                    for (int row = startRow; row <= endRow && !keepSeries; row++)
                    {
                        for (int col = startCol; col <= endCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];
                            if (cell.Type == CellValueType.IsNumeric && cell.DoubleValue > Threshold)
                            {
                                keepSeries = true;
                                break;
                            }
                        }
                    }

                    // Hide the series if no value exceeds the threshold
                    series.IsFiltered = !keepSeries;

                    Console.WriteLine($"Series {i + 1} ({rangeAddress}) " +
                                      (keepSeries ? "kept" : "filtered out"));
                }

                // Save the workbook with the filtered chart
                string outputPath = "FilteredChart.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving workbook: {ex.Message}");
                }

                // Export the chart to a PDF file (optional)
                string pdfPath = "FilteredChart.pdf";
                try
                {
                    chart.ToPdf(pdfPath);
                    Console.WriteLine($"Chart exported to PDF at {Path.GetFullPath(pdfPath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error exporting chart to PDF: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
