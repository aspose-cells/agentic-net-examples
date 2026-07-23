// Title: C# – Generate a Monthly‑Tick Timeline Chart and Save as High‑Resolution TIFF Using Aspose.Cells
// Description: This C# example builds a workbook, fills two columns with sequential dates and numeric values, adds a line chart whose category axis is configured as a time‑scale with months as the major unit and days as the minor unit, and renders the chart to a TIFF image using LZW compression at 300 dpi. The workbook is also saved for reference. The code demonstrates how to set monthly major ticks, control minor ticks, and produce print‑ready graphics; weekend dates can be omitted from the axis by applying a custom date filter.
// Keywords: Aspose.Cells | C# | .NET | timeline chart | monthly ticks | time scale axis | hide weekends | export TIFF | high resolution | LZW compression | 300 dpi | chart rendering | line chart
// Common Searches: Aspose.Cells create timeline chart with monthly axis ticks | Save Aspose.Cells chart as high‑resolution TIFF | C# time‑scale axis hide weekend dates Aspose.Cells | Configure major and minor units on Aspose.Cells chart axis | Export line chart to TIFF with LZW compression using Aspose.Cells
// Developer Intent: Create a line chart with a time‑scale axis that shows monthly major ticks, optionally excludes weekend dates, and export the chart as a high‑resolution TIFF image.
// Use Cases: Produce a printable timeline of project milestones with monthly gridlines for inclusion in PDF reports. | Generate a high‑quality chart image for marketing collateral where 300 dpi resolution and lossless compression are required. | Visualize daily sales data while displaying only month‑level markers and omitting weekend labels for clearer business analysis.
// AI Prompts: Show how to filter out weekend dates from the time‑scale axis in the Aspose.Cells chart. | Provide code to apply a custom date format to the category axis while keeping monthly major ticks. | Explain how to adjust TIFF compression and DPI settings in Aspose.Cells to balance file size and image quality.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;   // Required for ImageType and TiffCompression enums

namespace AsposeCellsExample
{
    // This C# example builds a workbook, fills two columns with sequential dates and numeric values, adds a line chart whose category axis is configured as a time‑scale with months as the major unit and days as the minor unit, and renders the chart to a TIFF image using LZW compression at 300 dpi. The workbook is also saved for reference. The code demonstrates how to set monthly major ticks, control minor ticks, and produce print‑ready graphics; weekend dates can be omitted from the axis by applying a custom date filter.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data with dates spanning several months (including weekends)
                sheet.Cells["A1"].PutValue("Date");
                sheet.Cells["B1"].PutValue("Value");
                DateTime startDate = new DateTime(2023, 1, 1);
                for (int i = 0; i < 60; i++) // 60 days ~ 2 months
                {
                    sheet.Cells[i + 1, 0].PutValue(startDate.AddDays(i));
                    sheet.Cells[i + 1, 1].PutValue(i * 10 + 50);
                }

                // Add a line chart that will display the dates on the category axis
                int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B61", true);
                chart.NSeries.CategoryData = "A2:A61";

                // Configure the category axis as a time scale with monthly major ticks
                Axis categoryAxis = chart.CategoryAxis;
                categoryAxis.CategoryType = CategoryType.TimeScale;
                categoryAxis.BaseUnitScale = TimeUnit.Months;   // Base unit = months
                categoryAxis.MajorUnitScale = TimeUnit.Months; // Major ticks = months
                categoryAxis.MinorUnitScale = TimeUnit.Days;   // Minor ticks = days

                // Prepare high‑quality TIFF image options
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Tiff,                     // Use Drawing.ImageType enum
                    TiffCompression = TiffCompression.CompressionLZW,
                    HorizontalResolution = 300,
                    VerticalResolution = 300
                };

                // Save the chart as a TIFF file
                string chartPath = "TimelineChart.tiff";
                try
                {
                    chart.ToImage(chartPath, imgOptions);
                    Console.WriteLine($"Chart image saved to {Path.GetFullPath(chartPath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save chart image: {ex.Message}");
                }

                // Optionally, save the workbook for reference
                string workbookPath = "TimelineWorkbook.xlsx";
                try
                {
                    workbook.Save(workbookPath);
                    Console.WriteLine($"Workbook saved to {Path.GetFullPath(workbookPath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
