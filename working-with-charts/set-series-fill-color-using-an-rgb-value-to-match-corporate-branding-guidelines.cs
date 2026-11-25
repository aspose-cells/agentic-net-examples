// Title: Apply a corporate RGB color to a chart series fill using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set chart series fill color with RGB values | change column chart series color programmatically with Aspose.Cells .NET | apply corporate branding color to Excel chart series using Aspose.Cells API | solid fill format for chart series in Aspose.Cells C# example | set series border color to match fill in Aspose.Cells chart | customize chart series appearance with RGB in Aspose.Cells for .NET
// Tags: solid fill format for chart series Aspose.Cells | RGB color for chart series Aspose.Cells | corporate branding colors Excel chart Aspose.Cells | chart series border styling Aspose.Cells | column chart series appearance Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;   // Required for FillType enum

namespace SetSeriesFillColorExample
{
    // The example creates a workbook, adds month and sales data, inserts a column chart, and then sets the first series' area fill to a solid corporate RGB color (34,120,200). It also applies the same color to the series border before saving the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and get the first worksheet.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // -----------------------------------------------------------------
                // 2. Populate sample data for the chart.
                // -----------------------------------------------------------------
                worksheet.Cells["A1"].PutValue("Month");
                worksheet.Cells["B1"].PutValue("Sales");

                string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
                double[] sales = { 12000, 15000, 13000, 17000, 16000 };

                for (int i = 0; i < months.Length; i++)
                {
                    worksheet.Cells[i + 1, 0].PutValue(months[i]);   // Column A
                    worksheet.Cells[i + 1, 1].PutValue(sales[i]);   // Column B
                }

                // -----------------------------------------------------------------
                // 3. Add a column chart to the worksheet.
                // -----------------------------------------------------------------
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the series.
                chart.NSeries.Add("B2:B6", true);
                chart.NSeries.CategoryData = "A2:A6";

                // -----------------------------------------------------------------
                // 4. Set the series fill color using a corporate RGB value.
                // -----------------------------------------------------------------
                if (chart.NSeries.Count > 0)
                {
                    Series firstSeries = chart.NSeries[0];

                    // Ensure the series area uses a solid fill.
                    firstSeries.Area.FillFormat.FillType = FillType.Solid;

                    // Apply the corporate RGB color.
                    Color corporateColor = Color.FromArgb(34, 120, 200);
                    firstSeries.Area.FillFormat.SolidFill.Color = corporateColor;

                    // Optionally set the border (useful for line/area charts).
                    firstSeries.Border.Color = corporateColor;
                }

                // -----------------------------------------------------------------
                // 5. Save the workbook to a file.
                // -----------------------------------------------------------------
                string outputPath = "SeriesFillColorDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
