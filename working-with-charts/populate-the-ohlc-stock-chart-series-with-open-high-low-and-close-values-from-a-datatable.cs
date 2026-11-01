// Title: Create an OHLC (Open‑High‑Low‑Close) stock chart from a DataTable using Aspose.Cells for .NET
// Common Searches: how to generate an OHLC stock chart from a DataTable with Aspose.Cells in C# | populate open high low close series in a StockOpenHighLowClose chart using Aspose.Cells | add series names to an OHLC chart created with Aspose.Cells .NET | write DataTable data to a worksheet and bind it to a StockOpenHighLowClose chart in C# | Aspose.Cells example for creating a financial OHLC chart from database data
// Tags: Aspose.Cells create OHLC chart | C# populate StockOpenHighLowClose series | Aspose.Cells bind DataTable to worksheet | Aspose.Cells set chart series names | Excel file generation with Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new Workbook, writes a DataTable containing Date, Open, High, Low, and Close columns to the first worksheet, adds a StockOpenHighLowClose chart referencing that range, sets the X‑axis to the Date column, adds four series for Open, High, Low, and Close with proper names, and saves the workbook as OHLC_Chart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Retrieve OHLC data (replace with actual data source as needed)
            DataTable dt = GetOhlcData();

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a meaningful name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "OHLCData";

            // Write header row
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Open");
            sheet.Cells["C1"].PutValue("High");
            sheet.Cells["D1"].PutValue("Low");
            sheet.Cells["E1"].PutValue("Close");

            // Populate worksheet with DataTable contents starting at row 2 (zero‑based index)
            int row = 1; // zero‑based index for the second row
            foreach (DataRow dr in dt.Rows)
            {
                sheet.Cells[row, 0].PutValue(dr["Date"]);
                sheet.Cells[row, 1].PutValue(dr["Open"]);
                sheet.Cells[row, 2].PutValue(dr["High"]);
                sheet.Cells[row, 3].PutValue(dr["Low"]);
                sheet.Cells[row, 4].PutValue(dr["Close"]);
                row++;
            }

            // Determine the range of data rows (Excel rows are 1‑based)
            string firstDataRow = "2";
            string lastDataRow = row.ToString(); // row already points to the last data row (1‑based)

            // Add a Stock (OHLC) chart below the data table
            int chartIndex = sheet.Charts.Add(ChartType.StockOpenHighLowClose, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "OHLC Stock Chart";

            // X‑axis (date) range
            string xRange = $"A{firstDataRow}:A{lastDataRow}";

            // Add series for Open, High, Low, Close in one call
            chart.NSeries.Add($"B{firstDataRow}:E{lastDataRow}", true);
            chart.NSeries.CategoryData = xRange;

            // Assign series names
            chart.NSeries[0].Name = "Open";
            chart.NSeries[1].Name = "High";
            chart.NSeries[2].Name = "Low";
            chart.NSeries[3].Name = "Close";

            // Save the workbook with the chart
            string outputPath = "OHLC_Chart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Example method that creates a sample DataTable with OHLC data
    static DataTable GetOhlcData()
    {
        DataTable dt = new DataTable();
        dt.Columns.Add("Date", typeof(DateTime));
        dt.Columns.Add("Open", typeof(double));
        dt.Columns.Add("High", typeof(double));
        dt.Columns.Add("Low", typeof(double));
        dt.Columns.Add("Close", typeof(double));

        // Sample rows – replace with real data retrieval logic
        dt.Rows.Add(new DateTime(2023, 1, 2), 100.0, 105.0, 95.0, 102.0);
        dt.Rows.Add(new DateTime(2023, 1, 3), 102.0, 108.0, 101.0, 107.0);
        dt.Rows.Add(new DateTime(2023, 1, 4), 107.0, 110.0, 106.0, 109.0);
        dt.Rows.Add(new DateTime(2023, 1, 5), 109.0, 112.0, 108.0, 111.0);

        return dt;
    }
}
