// Title: Create a column chart with Aspose.Cells for .NET and automatically resize the legend based on its content
// Common Searches: Aspose.Cells C# set legend size then auto‑adjust based on label length | programmatically calculate legend width for column chart in .NET Excel | dynamic legend resizing in Aspose.Cells chart using C# | how to auto size chart legend according to category names with Aspose.Cells | adjust Aspose.Cells chart legend dimensions after adding data series
// Tags: Aspose.Cells column chart legend auto‑size | C# calculate legend dimensions Aspose.Cells | dynamic legend resizing Excel chart .NET | chart legend size adjustment based on text Aspose.Cells | programmatic legend width height setting Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a new workbook, fills it with sample data, adds a column chart, sets an initial fixed legend size, then computes the longest category label and total text height to programmatically adjust the legend's width and height before saving the file as ChartWithDynamicLegend.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Alpha");
            sheet.Cells["B2"].PutValue(15);
            sheet.Cells["A3"].PutValue("Beta");
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["A4"].PutValue("Gamma");
            sheet.Cells["B4"].PutValue(45);

            // Add a column chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            // Category data line removed – Aspose.Cells can infer it from the first column.

            // ---- Legend: fixed size ----
            chart.Legend.Position = LegendPositionType.Right; // place legend on the right
            chart.Legend.Width = 100;   // initial width (points)
            chart.Legend.Height = 200;  // initial height (points)

            // ---- Programmatically adjust legend size based on its content ----
            const double charWidthApprox = 7; // rough width per character at default font size
            const double linePadding = 4;     // extra space between lines

            // Determine the longest category label (Alpha, Beta, Gamma)
            double maxTextWidth = 0;
            double totalTextHeight = 0;
            for (int row = 2; row <= 4; row++)
            {
                string text = sheet.Cells[row, 0].StringValue; // column A
                double entryWidth = text.Length * charWidthApprox;
                if (entryWidth > maxTextWidth)
                    maxTextWidth = entryWidth;

                totalTextHeight += chart.Legend.Font.Size + linePadding;
            }

            // Add some extra padding around the legend content
            chart.Legend.Width = (int)(maxTextWidth + 20);   // 20 points padding
            chart.Legend.Height = (int)(totalTextHeight + 20); // 20 points padding

            // Save the workbook
            string outputPath = "ChartWithDynamicLegend.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
