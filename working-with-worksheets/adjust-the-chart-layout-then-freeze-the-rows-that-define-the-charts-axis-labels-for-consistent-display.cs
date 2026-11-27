// Title: Adjust a column chart’s plot area and freeze axis label rows in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set custom plot area width and height for a column chart | Freeze top rows that contain chart axis labels with Aspose.Cells .NET | Move chart title and offset plot area in an Aspose.Cells column chart | Disable automatic chart sizing and enable inner mode in Aspose.Cells C# | Create a column chart and freeze rows 1‑6 using Aspose.Cells for .NET
// Tags: set plot area size Aspose.Cells | freeze panes rows Aspose.Cells | column chart layout Aspose.Cells | chart title positioning Aspose.Cells | disable automatic chart sizing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartLayoutAndFreezeDemo
{
    // The example creates a workbook, adds sample data, inserts a column chart, disables automatic sizing, defines explicit plot area dimensions and offsets, positions the chart title, freezes the first six rows that hold the axis labels, and saves the file as an XLSX workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the output workbook
            string outputPath = "ChartLayoutAndFreezeDemo.xlsx";

            try
            {
                // -------------------------------------------------
                // 1. Create a new workbook and get the first sheet
                // -------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // 2. Populate sample data that will be used for the chart
                // -------------------------------------------------
                // Header row (will be frozen)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                // Data rows (these rows contain the axis labels)
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B4"].PutValue(180);
                sheet.Cells["A5"].PutValue("Apr");
                sheet.Cells["B5"].PutValue(200);
                sheet.Cells["A6"].PutValue("May");
                sheet.Cells["B6"].PutValue(170);

                // -------------------------------------------------
                // 3. Add a column chart that uses the data above
                // -------------------------------------------------
                // Chart will be placed starting at row 8 (index 7) to leave space for frozen rows
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B6", true);          // Values
                chart.NSeries.CategoryData = "A2:A6";     // Categories (axis labels)

                // -------------------------------------------------
                // 4. Adjust the chart layout
                // -------------------------------------------------
                // Disable automatic sizing so we can control the exact dimensions
                chart.PlotArea.IsAutomaticSize = false;

                // Use inner mode so the size calculations include tick marks and axis labels
                chart.PlotArea.IsInnerMode = true;

                // Set explicit size for the plot area (values are in points; 1 point = 1/72 inch)
                chart.PlotArea.Width = 4000;   // Width of the plot area
                chart.PlotArea.Height = 2500;  // Height of the plot area

                // Optionally, shift the plot area a little to the right/down within the chart
                chart.PlotArea.X = 150;   // Horizontal offset from the chart's left edge
                chart.PlotArea.Y = 100;   // Vertical offset from the chart's top edge

                // Ensure the chart title is visible and positioned nicely
                chart.Title.Text = "Monthly Sales";
                chart.Title.YRatioToChart = 0.05; // 5% from the top of the chart area

                // -------------------------------------------------
                // 5. Freeze the rows that contain the axis labels
                // -------------------------------------------------
                // Freeze the first 6 rows (rows 0‑5). The overload requires four parameters:
                // row, column, freezableRows, freezableColumns.
                sheet.FreezePanes(6, 0, 6, 0);

                // -------------------------------------------------
                // 6. Save the workbook
                // -------------------------------------------------
                // Ensure the directory for the output file exists
                string directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved to '{outputPath}'.");
                Console.WriteLine("Chart layout adjusted and rows 1‑6 frozen for consistent axis label display.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
