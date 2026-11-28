// Title: Apply a built‑in monochromatic color palette and standardize fonts for a column chart using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# apply built‑in chart color palette to column chart | Change chart series colors to monochrome using Aspose.Cells .NET | Set chart title and axis font styles with Aspose.Cells in C# | Create a column chart with standardized fonts using Aspose.Cells | How to use ChangeColors method in Aspose.Cells for .NET
// Tags: Aspose.Cells ChangeColors method C# | built-in chart color palette Aspose.Cells | standardize chart fonts Aspose.Cells | column chart styling Aspose.Cells .NET | Excel chart theming Aspose.Cells | apply chart theme Aspose.Cells C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ApplyChartThemeDemo
{
    // The example creates a workbook, adds sample sales data, generates a column chart, applies the built‑in MonochromaticPalette6 via ChangeColors, sets the chart title, legend, and axis titles to Calibri fonts, and saves the file as ChartWithBuiltInTheme.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 2. Populate sample data for the chart
                cells["A1"].PutValue("Month");
                cells["B1"].PutValue("Sales");
                cells["A2"].PutValue("Jan");
                cells["A3"].PutValue("Feb");
                cells["A4"].PutValue("Mar");
                cells["A5"].PutValue("Apr");
                cells["B2"].PutValue(120);
                cells["B3"].PutValue(150);
                cells["B4"].PutValue(180);
                cells["B5"].PutValue(200);

                // 3. Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // 4. Set the data range for the chart
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // 5. Apply a built‑in color palette to the chart series
                SeriesCollection seriesColl = chart.NSeries;
                seriesColl.ChangeColors(ChartColorPaletteType.MonochromaticPalette6);

                // 6. Standardize fonts for chart elements
                // Title
                chart.Title.Text = "Quarterly Sales";
                chart.Title.Font.Name = "Calibri";
                chart.Title.Font.Size = 14;
                chart.Title.Font.IsBold = true;

                // Legend
                chart.Legend.Font.Name = "Calibri";
                chart.Legend.Font.Size = 11;

                // Axis titles
                chart.CategoryAxis.Title.Text = "Month";
                chart.CategoryAxis.Title.Font.Name = "Calibri";
                chart.CategoryAxis.Title.Font.Size = 12;
                chart.CategoryAxis.Title.Font.IsBold = true;

                chart.ValueAxis.Title.Text = "Sales (Units)";
                chart.ValueAxis.Title.Font.Name = "Calibri";
                chart.ValueAxis.Title.Font.Size = 12;
                chart.ValueAxis.Title.Font.IsBold = true;

                // 7. Save the workbook
                string outputPath = "ChartWithBuiltInTheme.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while applying the chart theme:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
