// Title: Add a line chart with cell‑linked data labels to every worksheet in an Aspose.Cells .NET workbook
// AI Prompts: Create a line chart on each worksheet and configure its data labels to display values from a separate label column. | Loop through all worksheets, fill them with sample data, and link the chart series data labels to the cells in column C.
// Common Searches: Aspose.Cells how to add a line chart to each worksheet programmatically | C# link chart data labels to a cell range using Aspose.Cells | batch generate charts with custom labels from cells in Aspose.Cells .NET | set data label source for line chart series in Aspose.Cells workbook
// Tags: line chart per worksheet Aspose.Cells | link chart data labels to cell range Aspose.Cells | batch chart generation .NET workbook | custom data labels from column C Aspose.Cells | populate worksheets with sample data C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook with three worksheets, adds sample data to each, inserts a line chart on every sheet, and links the chart's data labels to the corresponding cells in column C.
class BatchLineChartWithCellLabels
{
    static void Main()
    {
        try
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Add two more worksheets so we have three in total
            workbook.Worksheets.Add();
            workbook.Worksheets.Add();

            // Prepare sample data in each worksheet
            for (int wsIndex = 0; wsIndex < 3; wsIndex++)
            {
                Worksheet sheet = workbook.Worksheets[wsIndex];
                sheet.Name = $"Sheet{wsIndex + 1}";

                // Header
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["C1"].PutValue("Label");

                // Sample rows
                for (int i = 2; i <= 6; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                    sheet.Cells[$"B{i}"].PutValue(i * 10);               // numeric value
                    sheet.Cells[$"C{i}"].PutValue($"Label {i - 1}");    // custom label to link
                }

                // Determine the last row with data (excluding header)
                int lastRow = sheet.Cells.MaxDataRow; // zero‑based index
                int dataStartRow = 2;                  // first data row (1‑based)
                int dataEndRow = lastRow + 1;          // convert to 1‑based row number

                // Add a line chart to the current worksheet
                int chartIdx = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];

                // Set data source for the series and categories
                chart.NSeries.Add($"B{dataStartRow}:B{dataEndRow}", true);
                chart.NSeries.CategoryData = $"A{dataStartRow}:A{dataEndRow}";

                // Configure the series to use cell values as data labels
                Series series = chart.NSeries[0];
                series.DataLabels.ShowCellRange = true;                                 // show cell range as label
                series.DataLabels.LinkedSource = $"C{dataStartRow}:C{dataEndRow}";      // link to label cells
                series.DataLabels.ShowValue = false;                                   // hide default numeric value
            }

            // Save the workbook
            workbook.Save("BatchLineCharts.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
