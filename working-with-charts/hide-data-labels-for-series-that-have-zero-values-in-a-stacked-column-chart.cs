// Title: How to hide data labels for series that contain only zero values in a stacked column chart using Aspose.Cells for .NET
// Common Searches: Aspose.Cells hide data labels for series with all zero values in stacked column chart C# | C# conditional chart label visibility based on zero values with Aspose.Cells | Remove data labels from zero-valued series in an Excel stacked column chart using Aspose.Cells | Programmatically suppress data labels for empty series in Aspose.Cells charts | Aspose.Cells hide series data labels when points are zero in .NET
// Tags: hide zero-valued series data labels Aspose.Cells | stacked column chart label visibility .NET | conditional data label display C# Aspose.Cells | chart series zero value handling Aspose.Cells | Excel chart data label suppression for empty series

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace HideZeroSeriesDataLabels
{
    // The example creates a workbook with sample data that includes zero values, adds a stacked column chart, enables data labels for all series, then iterates each series to detect if every point is zero and hides that series' data labels by setting the DataLabels.IsDeleted property, finally saving the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data with some zero values
                // Categories (X axis)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["A5"].PutValue("Q4");

                // Series 1
                sheet.Cells["B1"].PutValue("Product A");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(0);   // zero value
                sheet.Cells["B4"].PutValue(150);
                sheet.Cells["B5"].PutValue(130);

                // Series 2 (all zeros – should have its data labels hidden)
                sheet.Cells["C1"].PutValue("Product B");
                sheet.Cells["C2"].PutValue(0);
                sheet.Cells["C3"].PutValue(0);
                sheet.Cells["C4"].PutValue(0);
                sheet.Cells["C5"].PutValue(0);

                // Series 3
                sheet.Cells["D1"].PutValue("Product C");
                sheet.Cells["D2"].PutValue(80);
                sheet.Cells["D3"].PutValue(95);
                sheet.Cells["D4"].PutValue(0);   // zero value
                sheet.Cells["D5"].PutValue(110);

                // Add a stacked column chart (use ColumnStacked enum value)
                int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart (all series)
                chart.NSeries.Add("B2:D5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Enable data labels for all series initially
                foreach (Series series in chart.NSeries)
                {
                    series.DataLabels.ShowValue = true;
                    series.DataLabels.IsDeleted = false; // ensure labels are visible
                }

                // Hide data labels for any series whose all points are zero
                foreach (Series series in chart.NSeries)
                {
                    bool allZero = true;
                    foreach (ChartPoint point in series.Points)
                    {
                        // ChartPoint.YValue returns object; convert safely to double
                        if (point.YValue != null)
                        {
                            double yVal = Convert.ToDouble(point.YValue);
                            if (Math.Abs(yVal) > double.Epsilon)
                            {
                                allZero = false;
                                break;
                            }
                        }
                    }

                    if (allZero)
                    {
                        series.DataLabels.IsDeleted = true; // hide labels for this series
                    }
                }

                // Save the workbook
                string outputPath = "StackedColumn_HideZeroSeriesDataLabels.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
