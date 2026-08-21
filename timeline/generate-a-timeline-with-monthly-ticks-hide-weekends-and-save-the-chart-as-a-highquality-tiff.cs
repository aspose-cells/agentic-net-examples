// Title: C# – Create a PivotTable Timeline with Monthly Ticks and Export the Chart as a High‑Resolution TIFF using Aspose.Cells
// Description: This example shows how to use Aspose.Cells for .NET to generate a workbook, populate it with date/value rows, build a PivotTable, attach a Timeline control, configure a line chart with a time‑scale axis (monthly major ticks, daily minor ticks, weekends hidden), and render the chart to a 300 dpi LZW‑compressed TIFF file while also saving the workbook.
// Keywords: Aspose.Cells timeline chart | C# PivotTable timeline | monthly major ticks Aspose.Cells | hide weekends time scale chart | export chart to TIFF | high resolution TIFF Aspose.Cells | LZW compression | 300 dpi chart image | .NET chart rendering | Aspose.Cells example
// Common Searches: Aspose.Cells add timeline to pivot table C# | set monthly major tick marks on time scale axis Aspose.Cells | hide weekends on chart axis Aspose.Cells | save Aspose.Cells chart as high quality TIFF | export line chart to TIFF with LZW compression
// Developer Intent: Generate a timeline linked to a PivotTable, configure monthly tick marks on a time‑scale axis, hide weekend dates, and save the resulting chart as a high‑resolution TIFF image.
// Use Cases: Produce monthly reporting charts where each major tick represents one month. | Enable interactive month‑by‑month filtering through a Timeline control attached to a PivotTable. | Create print‑ready graphics by exporting charts to 300 dpi LZW‑compressed TIFF files.
// AI Prompts: Write C# code with Aspose.Cells to add a Timeline to a PivotTable and set the category axis to show monthly major ticks while hiding weekends. | Show how to render an Aspose.Cells line chart to a 300 dpi TIFF image using LZW compression. | Explain how to configure minor tick units for daily intervals on a time‑scale axis in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;   // Required for ImageType and TiffCompression enums

namespace AsposeCellsTimelineChartDemo
{
    // This example shows how to use Aspose.Cells for .NET to generate a workbook, populate it with date/value rows, build a PivotTable, attach a Timeline control, configure a line chart with a time‑scale axis (monthly major ticks, daily minor ticks, weekends hidden), and render the chart to a 300 dpi LZW‑compressed TIFF file while also saving the workbook.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // 1. Populate worksheet with sample date/value data
                // -------------------------------------------------
                // Header
                sheet.Cells["A1"].PutValue("Date");
                sheet.Cells["B1"].PutValue("Value");

                // Sample data spanning several months (including weekends)
                DateTime startDate = new DateTime(2023, 1, 1);
                for (int i = 0; i < 60; i++) // 60 days ~ 2 months
                {
                    sheet.Cells[i + 1, 0].PutValue(startDate.AddDays(i));
                    sheet.Cells[i + 1, 1].PutValue(i * 10 + 100);
                }

                // -------------------------------------------------
                // 2. Create a PivotTable (required as data source for Timeline)
                // -------------------------------------------------
                PivotTableCollection pivots = sheet.PivotTables;
                int pivotIndex = pivots.Add("A1:B61", "D1", "SalesPivot");
                PivotTable pivot = pivots[pivotIndex];
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Value");

                // Refresh pivot cache and calculate data (use correct API)
                pivot.RefreshData();
                pivot.CalculateData();

                // -------------------------------------------------
                // 3. Add a Timeline linked to the PivotTable
                // -------------------------------------------------
                // Place the Timeline at row 10, column 5 (E10)
                int timelineIndex = sheet.Timelines.Add(pivot, 9, 4, "Date");
                Timeline timeline = sheet.Timelines[timelineIndex];
                timeline.Caption = "Monthly Timeline";

                // -------------------------------------------------
                // 4. Create a Line chart that uses the same date range
                // -------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Line, 15, 0, 30, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set data source for the chart
                chart.NSeries.Add("B2:B61", true);
                chart.NSeries.CategoryData = "A2:A61";

                // Configure the category axis as a time scale with monthly major ticks
                Axis categoryAxis = chart.CategoryAxis;
                categoryAxis.CategoryType = CategoryType.TimeScale;
                categoryAxis.BaseUnitScale = TimeUnit.Months;   // Base unit = months
                categoryAxis.MajorUnitScale = TimeUnit.Months; // Major ticks = months
                categoryAxis.MajorUnit = 1;                    // One month per major tick
                categoryAxis.MinorUnitScale = TimeUnit.Days;   // Minor ticks = days
                categoryAxis.MinorUnit = 1;                    // One day per minor tick

                // -------------------------------------------------
                // 5. Save the chart as a high‑quality TIFF image
                // -------------------------------------------------
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Tiff,
                    TiffCompression = TiffCompression.CompressionLZW,
                    HorizontalResolution = 300, // High‑quality DPI
                    VerticalResolution = 300
                };

                string tiffPath = "ChartTimelineHighQuality.tiff";
                chart.ToImage(tiffPath, imgOptions);

                // -------------------------------------------------
                // 6. Save the workbook (optional, to keep the timeline)
                // -------------------------------------------------
                string workbookPath = "WorkbookWithTimeline.xlsx";
                workbook.Save(workbookPath);

                Console.WriteLine("Chart saved as high‑quality TIFF: " + tiffPath);
                Console.WriteLine("Workbook saved with timeline: " + workbookPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
