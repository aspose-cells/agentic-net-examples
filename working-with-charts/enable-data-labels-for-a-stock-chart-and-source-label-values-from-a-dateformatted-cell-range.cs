// Title: Add Date‑Based Data Labels to a Stock OHLC Chart with Aspose.Cells for .NET
// Description: Creates a workbook, fills columns A‑E with dates and OHLC values, applies a custom date format, inserts a StockOpenHighLowClose chart, and configures the series to show data labels that pull their text from the formatted date cells (A2:A6) while preserving the cell's date format, then saves the file.
// Keywords: Aspose.Cells stock chart | C# data labels from cell range | date formatted chart labels | ShowCellRange Aspose.Cells | LinkedSource chart labels | NumberFormatLinked property | Open‑High‑Low‑Close chart .NET | Excel chart data labels C#
// Common Searches: Aspose.Cells display dates as data labels on stock chart | Enable data labels from a cell range in .NET | Link chart labels to worksheet cells Aspose.Cells | Preserve cell date format in chart data labels | Create OHLC chart with linked date labels C#
// Developer Intent: Create a StockOpenHighLowClose chart and configure its data labels to show dates taken from a formatted cell range.
// Use Cases: Financial reporting workbook where each OHLC point is annotated with its trading date. | Dynamic charts that automatically reflect changes to the date column without code modifications. | Exporting Excel files with consistent date formatting on chart labels for regulatory compliance.
// AI Prompts: Generate C# code using Aspose.Cells to build a StockOpenHighLowClose chart, enable data labels, link them to a date column, and retain the cell's date format. | Explain the roles of ShowCellRange, LinkedSource, and NumberFormatLinked when setting up series data labels in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills columns A‑E with dates and OHLC values, applies a custom date format, inserts a StockOpenHighLowClose chart, and configures the series to show data labels that pull their text from the formatted date cells (A2:A6) while preserving the cell's date format, then saves the file.
class StockChartDataLabelsDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: dates in column A, OHLC values in columns B‑E
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Open");
            sheet.Cells["C1"].PutValue("High");
            sheet.Cells["D1"].PutValue("Low");
            sheet.Cells["E1"].PutValue("Close");

            DateTime startDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(startDate.AddDays(i)); // Date
                sheet.Cells[i + 1, 1].PutValue(100 + i * 5);        // Open
                sheet.Cells[i + 1, 2].PutValue(110 + i * 5);        // High
                sheet.Cells[i + 1, 3].PutValue(90 + i * 5);         // Low
                sheet.Cells[i + 1, 4].PutValue(105 + i * 5);        // Close
            }

            // Apply a date format to the date column (A2:A6)
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "m/d/yyyy";

            for (int row = 2; row <= 6; row++)
            {
                sheet.Cells[row, 0].SetStyle(dateStyle);
            }

            // Add a Stock chart (Open‑High‑Low‑Close)
            int chartIndex = sheet.Charts.Add(ChartType.StockOpenHighLowClose, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set series data (Open‑High‑Low‑Close) and category (dates)
            chart.NSeries.Add("B2:E6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Enable data labels and link them to the date cells
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;          // show the OHLC value
            series.DataLabels.ShowCellRange = true;      // use cell range for label text
            series.DataLabels.LinkedSource = "A2:A6";    // link to the date cells
            series.DataLabels.NumberFormatLinked = true; // keep the date format from cells

            // Save the workbook
            string outputPath = "StockChartWithDateLabels.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
