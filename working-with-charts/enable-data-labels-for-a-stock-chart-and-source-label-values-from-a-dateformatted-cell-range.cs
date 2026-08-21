// Title: Aspose.Cells for .NET – Add Date‑Formatted Data Labels to a Stock OHLC Chart (C#)
// Description: Demonstrates how to create a workbook, populate it with OHLC stock data, format a column as dates, and link those date cells to the chart’s data labels. The example shows enabling data labels, setting ShowCellRange, assigning a LinkedSource range, hiding default values, and saving the file while ensuring the output folder exists.
// Keywords: Aspose.Cells | C# stock chart | OHLC chart data labels | date formatted labels | ShowCellRange | LinkedSource | custom chart labels | Aspose.Cells example | Excel chart automation | date style in cells
// Common Searches: Aspose.Cells add data labels to stock chart | C# link chart labels to cell range | format chart label dates Aspose.Cells | ShowCellRange property example | stock Open High Low Close chart with custom labels | Aspose.Cells date style for chart labels
// Developer Intent: Add custom date‑formatted data labels to a Stock Open‑High‑Low‑Close chart by linking them to a cell range.
// Use Cases: Generate a financial workbook with OHLC data and display trade dates as data labels. | Apply a custom date format to label cells and hide numeric values on the chart. | Automate chart creation and labeling in server‑side .NET applications.
// AI Prompts: Write C# code using Aspose.Cells to create a StockOpenHighLowClose chart and use a date‑formatted cell range for its data labels. | Explain how Series.DataLabels.ShowCellRange, LinkedSource, and ShowValue work together to display custom labels on a stock chart. | Provide troubleshooting steps when linked source cells do not appear as data labels in an Aspose.Cells chart.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to create a workbook, populate it with OHLC stock data, format a column as dates, and link those date cells to the chart’s data labels. The example shows enabling data labels, setting ShowCellRange, assigning a LinkedSource range, hiding default values, and saving the file while ensuring the output folder exists.
class StockChartDataLabelsDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare sample data for a stock chart
            // Column A: Date (category)
            // Columns B‑E: Open, High, Low, Close values
            // Column F: Date values that will be used as data‑label text
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Open");
            sheet.Cells["C1"].PutValue("High");
            sheet.Cells["D1"].PutValue("Low");
            sheet.Cells["E1"].PutValue("Close");
            sheet.Cells["F1"].PutValue("LabelDate");

            DateTime dt1 = new DateTime(2023, 1, 1);
            sheet.Cells["A2"].PutValue(dt1);
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["C2"].PutValue(110);
            sheet.Cells["D2"].PutValue(95);
            sheet.Cells["E2"].PutValue(105);
            sheet.Cells["F2"].PutValue(dt1);

            DateTime dt2 = new DateTime(2023, 1, 2);
            sheet.Cells["A3"].PutValue(dt2);
            sheet.Cells["B3"].PutValue(106);
            sheet.Cells["C3"].PutValue(115);
            sheet.Cells["D3"].PutValue(102);
            sheet.Cells["E3"].PutValue(112);
            sheet.Cells["F3"].PutValue(dt2);

            // Apply a date format to the label cells (F2:F3)
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "mm-dd-yyyy";

            // Define a StyleFlag to apply all style attributes
            StyleFlag flag = new StyleFlag { All = true };

            // Create a range for F2:F3 and apply the style
            AsposeRange labelRange = sheet.Cells.CreateRange("F2", "F3");
            labelRange.ApplyStyle(dateStyle, flag); // Apply style with flag

            // Add a Stock Open‑High‑Low‑Close chart
            int chartIndex = sheet.Charts.Add(ChartType.StockOpenHighLowClose, 5, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the series data (Open‑High‑Low‑Close) and category (Date)
            chart.NSeries.Add("B2:E3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Enable data labels and source their text from the date‑formatted range
            Series series = chart.NSeries[0];
            series.DataLabels.ShowCellRange = true;          // Use cell range for labels
            series.DataLabels.LinkedSource = "F2:F3";        // Cells containing the label dates
            series.DataLabels.ShowValue = false;            // Hide the default numeric values

            // Save the workbook
            string outputPath = "StockChartDataLabelsDemo.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
