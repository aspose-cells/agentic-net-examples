// Title: Aspose.Cells .NET – Create an OHLC (Open‑High‑Low‑Close) Stock Chart from a DataTable
// Description: This C# example shows how to build a DataTable with Date, Open, High, Low, and Close columns, write it to the first worksheet of a new Workbook, add a StockOpenHighLowClose chart, bind each OHLC series to the appropriate range, set the date column as the category axis, optionally name the series, and save the result as an Excel file.
// Keywords: Aspose.Cells | .NET | C# | OHLC chart | Open High Low Close | StockOpenHighLowClose | DataTable to chart | Excel chart generation | financial charting | global finance
// Common Searches: Aspose.Cells create OHLC chart C# | bind DataTable to StockOpenHighLowClose chart | C# OHLC chart from Excel data | set date category axis in Aspose.Cells chart | add series names to OHLC chart Aspose.Cells
// Developer Intent: Generate an Excel workbook that contains an Open‑High‑Low‑Close stock chart populated directly from a DataTable.
// Use Cases: Produce daily stock price visualizations for financial reports by converting market data tables into OHLC charts. | Automate the creation of separate OHLC charts for multiple securities in a batch process, exporting each to its own Excel file. | Provide analysts with ready‑to‑use Excel workbooks that include OHLC charts for further analysis or presentation.
// AI Prompts: Show me how to import a DataTable into a worksheet and bind it to an OHLC chart using Aspose.Cells for .NET. | Give a concise example of assigning series names and date categories to a StockOpenHighLowClose chart. | Explain how to customize colors and line styles of an OHLC chart after adding the series in Aspose.Cells.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsOHLCExample
{
    // This C# example shows how to build a DataTable with Date, Open, High, Low, and Close columns, write it to the first worksheet of a new Workbook, add a StockOpenHighLowClose chart, bind each OHLC series to the appropriate range, set the date column as the category axis, optionally name the series, and save the result as an Excel file.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- 1. Prepare a DataTable with OHLC data ----------
                DataTable ohlcTable = new DataTable("OHLC");
                ohlcTable.Columns.Add("Date", typeof(DateTime));
                ohlcTable.Columns.Add("Open", typeof(double));
                ohlcTable.Columns.Add("High", typeof(double));
                ohlcTable.Columns.Add("Low", typeof(double));
                ohlcTable.Columns.Add("Close", typeof(double));

                // Sample rows
                ohlcTable.Rows.Add(new DateTime(2023, 1, 2), 100.5, 105.2, 99.8, 104.0);
                ohlcTable.Rows.Add(new DateTime(2023, 1, 3), 104.0, 108.5, 103.2, 107.1);
                ohlcTable.Rows.Add(new DateTime(2023, 1, 4), 107.1, 110.0, 106.5, 109.3);
                ohlcTable.Rows.Add(new DateTime(2023, 1, 5), 109.3, 112.4, 108.7, 111.0);
                ohlcTable.Rows.Add(new DateTime(2023, 1, 6), 111.0, 113.8, 110.2, 112.5);

                // ---------- 2. Create a workbook ----------
                Workbook workbook = new Workbook();                     // create workbook
                Worksheet sheet = workbook.Worksheets[0];              // get first worksheet

                // Manually import DataTable into worksheet (avoids ImportDataTable overload issues)
                // Write column headers
                for (int col = 0; col < ohlcTable.Columns.Count; col++)
                {
                    sheet.Cells[0, col].PutValue(ohlcTable.Columns[col].ColumnName);
                }

                // Write rows
                for (int row = 0; row < ohlcTable.Rows.Count; row++)
                {
                    for (int col = 0; col < ohlcTable.Columns.Count; col++)
                    {
                        sheet.Cells[row + 1, col].PutValue(ohlcTable.Rows[row][col]);
                    }
                }

                // ---------- 3. Add an OHLC (Open‑High‑Low‑Close) stock chart ----------
                // Chart type for OHLC is StockOpenHighLowClose
                int chartIndex = sheet.Charts.Add(ChartType.StockOpenHighLowClose, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // ---------- 4. Define the data ranges ----------
                // Open series
                chart.NSeries.Add("=Sheet1!$B$2:$B$6", true);
                // High series
                chart.NSeries.Add("=Sheet1!$C$2:$C$6", true);
                // Low series
                chart.NSeries.Add("=Sheet1!$D$2:$D$6", true);
                // Close series
                chart.NSeries.Add("=Sheet1!$E$2:$E$6", true);

                // Category (X‑axis) data – dates
                chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$6";

                // ---------- 5. (Optional) Set series names ----------
                chart.NSeries[0].Name = "Open";
                chart.NSeries[1].Name = "High";
                chart.NSeries[2].Name = "Low";
                chart.NSeries[3].Name = "Close";

                // ---------- 6. Save the workbook ----------
                string outputPath = "OHLC_StockChart.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
