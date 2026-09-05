// Title: How to add date‑formatted data labels to a Stock Open‑High‑Low‑Close chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a StockOpenHighLowClose chart and binds its data labels to a range of cells containing dates, keeping the original number format. | Demonstrate enabling ShowCellRange, setting LinkedSource, and applying custom font styling for series data labels in an Aspose.Cells chart.
// Common Searches: Aspose.Cells C# link chart data labels to a date column | Create StockOpenHighLowClose chart with custom label text from cells | Display dates as data labels in an Aspose.Cells stock chart | Set ShowCellRange and LinkedSource for series labels in Aspose.Cells .NET | Apply date number format to chart label range using Aspose.Cells
// Tags: Aspose.Cells chart series data label linking | C# date cell range as chart label text | Enable ShowCellRange for Aspose.Cells charts | StockOpenHighLowClose chart with custom labels | Apply number format to chart data labels Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsStockChartDataLabels
{
    // Creates a workbook, writes dates and open prices, copies dates to a label column with short date formatting, adds a StockOpenHighLowClose chart, enables ShowCellRange, sets LinkedSource to the date column, links the number format, customizes label font color, and saves the file as StockChartWithDateLabels.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Populate sample data
                // -------------------------------------------------
                // Header row
                sheet.Cells["A1"].PutValue("Date");
                sheet.Cells["B1"].PutValue("Open");
                sheet.Cells["C1"].PutValue("LabelDate");

                // Sample dates and values
                sheet.Cells["A2"].PutValue(DateTime.Today);
                sheet.Cells["A3"].PutValue(DateTime.Today.AddDays(1));
                sheet.Cells["A4"].PutValue(DateTime.Today.AddDays(2));
                sheet.Cells["A5"].PutValue(DateTime.Today.AddDays(3));

                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(125);
                sheet.Cells["B4"].PutValue(123);
                sheet.Cells["B5"].PutValue(128);

                // Copy dates to column C (the range that will be used for data‑label text)
                sheet.Cells["C2"].PutValue(DateTime.Today);
                sheet.Cells["C3"].PutValue(DateTime.Today.AddDays(1));
                sheet.Cells["C4"].PutValue(DateTime.Today.AddDays(2));
                sheet.Cells["C5"].PutValue(DateTime.Today.AddDays(3));

                // Apply a date number format to columns A and C
                Style dateStyle = workbook.CreateStyle();
                dateStyle.Number = 14; // Built‑in short date format (e.g., m/d/yyyy)
                StyleFlag flag = new StyleFlag { NumberFormat = true };
                sheet.Cells.CreateRange("A2:A5").ApplyStyle(dateStyle, flag);
                sheet.Cells.CreateRange("C2:C5").ApplyStyle(dateStyle, flag);

                // -------------------------------------------------
                // Add a Stock chart (using Open‑High‑Low‑Close type; only Open series is provided)
                // -------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.StockOpenHighLowClose, 6, 0, 22, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the series (Open values) and categories (dates)
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // -------------------------------------------------
                // Configure data labels to show values from the date‑formatted range (C2:C5)
                // -------------------------------------------------
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;          // Show the numeric value (Open)
                series.DataLabels.ShowCellRange = true;      // Enable using a cell range for label text
                series.DataLabels.LinkedSource = "C2:C5";    // Use the formatted dates as label text
                series.DataLabels.NumberFormatLinked = true; // Keep number format linked to source cells
                series.DataLabels.Font.Color = Color.Blue;   // Optional: make label text blue

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                workbook.Save("StockChartWithDateLabels.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
