using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace OhlcChartExampleApp
{
    class OhlcChartExample
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                string sheetName = sheet.Name; // usually "Sheet1"

                // 2. Prepare a DataTable with Date, Open, High, Low, Close columns
                DataTable dt = new DataTable();
                dt.Columns.Add("Date", typeof(DateTime));
                dt.Columns.Add("Open", typeof(double));
                dt.Columns.Add("High", typeof(double));
                dt.Columns.Add("Low", typeof(double));
                dt.Columns.Add("Close", typeof(double));

                // Sample data rows
                dt.Rows.Add(new DateTime(2023, 1, 2), 100.5, 105.2, 99.8, 104.0);
                dt.Rows.Add(new DateTime(2023, 1, 3), 104.0, 108.5, 103.2, 107.1);
                dt.Rows.Add(new DateTime(2023, 1, 4), 107.1, 110.0, 106.5, 109.3);
                dt.Rows.Add(new DateTime(2023, 1, 5), 109.3, 112.4, 108.7, 111.0);
                dt.Rows.Add(new DateTime(2023, 1, 6), 111.0, 113.8, 110.2, 112.5);

                // 3. Import the DataTable into the worksheet manually
                int currentRow = 0;
                // Write column headers
                for (int c = 0; c < dt.Columns.Count; c++)
                    sheet.Cells[currentRow, c].PutValue(dt.Columns[c].ColumnName);
                currentRow++;

                // Write data rows
                foreach (DataRow dr in dt.Rows)
                {
                    for (int c = 0; c < dt.Columns.Count; c++)
                        sheet.Cells[currentRow, c].PutValue(dr[c]);
                    currentRow++;
                }

                // 4. Add a Stock OHLC chart to the worksheet
                // Parameters: chart type, top row, left column, bottom row, right column
                int chartIndex = sheet.Charts.Add(ChartType.StockOpenHighLowClose, 8, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // 5. Set the category (X) axis data to the Date column (A2:A{last})
                string catRange = $"='{sheetName}'!$A${2}:$A${dt.Rows.Count + 1}";
                chart.NSeries.CategoryData = catRange;

                // 6. Add series for Open, High, Low, Close values
                string openRange = $"='{sheetName}'!$B${2}:$B${dt.Rows.Count + 1}";
                string highRange = $"='{sheetName}'!$C${2}:$C${dt.Rows.Count + 1}";
                string lowRange = $"='{sheetName}'!$D${2}:$D${dt.Rows.Count + 1}";
                string closeRange = $"='{sheetName}'!$E${2}:$E${dt.Rows.Count + 1}";

                chart.NSeries.Add(openRange, true);   // Open series
                chart.NSeries.Add(highRange, false); // High series
                chart.NSeries.Add(lowRange, false);  // Low series
                chart.NSeries.Add(closeRange, false);// Close series

                // 7. Save the workbook
                string outputPath = "OhlcChart.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}