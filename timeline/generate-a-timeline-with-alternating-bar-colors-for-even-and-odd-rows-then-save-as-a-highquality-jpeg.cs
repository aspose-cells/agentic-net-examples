// Title: Create a timeline bar chart with alternating bar colors and save it as a high‑resolution JPEG in C# using Aspose.Cells
// Common Searches: c# asp.net how to apply alternating bar colors in an Aspose.Cells chart | aspocells export chart to jpeg with 300 dpi and maximum quality in .NET | generate a timeline bar chart with weekly dates using Aspose.Cells C# | set individual bar colors programmatically in Aspose.Cells chart series | save Aspose.Cells chart as high resolution JPEG image file
// Tags: bar chart point color customization Aspose.Cells | export chart as high resolution JPEG Aspose.Cells | chronological bar chart creation Aspose.Cells C# | configure image resolution and quality Aspose.Cells | set series point foreground color Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing; // For ImageType enum

namespace TimelineChartDemo
{
    // The example creates a new workbook, fills column A with weekly dates starting 2023‑01‑01 and column B with incremental values, adds a bar chart to act as a timeline, assigns blue to even‑indexed bars and red to odd‑indexed bars by setting each point's foreground color, configures ImageOrPrintOptions for 300 dpi and 100 % quality, renders the chart to a high‑resolution JPEG file, and saves the workbook as an XLSX file.
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

                // -------------------------------------------------
                // 1. Populate sample data (Date/Value)
                // -------------------------------------------------
                cells["A1"].PutValue("Date");
                cells["B1"].PutValue("Value");

                DateTime startDate = new DateTime(2023, 1, 1);
                for (int i = 0; i < 10; i++)
                {
                    cells[i + 2, 0].PutValue(startDate.AddDays(i * 7)); // weekly dates
                    cells[i + 2, 1].PutValue(50 + i * 10);            // arbitrary values
                }

                // -------------------------------------------------
                // 2. Add a Bar chart (acts as a timeline)
                // -------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Bar, 2, 3, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series (values) and categories (dates)
                chart.SetChartDataRange("A1:B11", true);

                // Optional: give the chart a title
                chart.Title.Text = "Timeline with Alternating Bar Colors";

                // -------------------------------------------------
                // 3. Apply alternating colors to each data point
                // -------------------------------------------------
                Color evenColor = Color.FromArgb(0, 112, 192); // blue
                Color oddColor = Color.FromArgb(192, 80, 77);  // red

                // Ensure there is at least one series before accessing it
                if (chart.NSeries.Count > 0)
                {
                    for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
                    {
                        ChartPoint point = chart.NSeries[0].Points[i];
                        point.Area.ForegroundColor = (i % 2 == 0) ? evenColor : oddColor;
                    }
                }

                // -------------------------------------------------
                // 4. Save the chart as a high‑quality JPEG image
                // -------------------------------------------------
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Jpeg,
                    HorizontalResolution = 300,
                    VerticalResolution = 300,
                    Quality = 100
                };

                string outputPath = "TimelineChart.jpg";

                // Ensure the directory exists (if a directory is specified)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Render chart to image file
                chart.ToImage(outputPath, imgOptions);
                Console.WriteLine($"Chart successfully saved to \"{outputPath}\" with high‑quality JPEG settings.");

                // (Optional) Save the workbook for reference
                string workbookPath = "TimelineChartWorkbook.xlsx";
                workbook.Save(workbookPath);
                Console.WriteLine($"Workbook saved to \"{workbookPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
