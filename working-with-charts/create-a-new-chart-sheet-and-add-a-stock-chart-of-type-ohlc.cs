// Title: Add a Chart Sheet with an OHLC Stock Chart using Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, fills a data worksheet with Date, Open, High, Low, and Close values, adds a separate chart sheet, inserts an Open‑High‑Low‑Close (OHLC) stock chart, links the series to columns B‑E, sets the X‑axis to column A, adds a title, and saves the file as an XLSX workbook.
// Keywords: Aspose.Cells chart sheet | OHLC stock chart C# | Aspose.Cells add chart sheet | Open‑High‑Low‑Close chart .NET | Excel OHLC chart Aspose
// Common Searches: Aspose.Cells create chart sheet OHLC | C# add OHLC stock chart to workbook | How to generate separate chart sheet with Aspose.Cells | OHLC chart example Aspose.Cells .NET | Set series range for OHLC chart Aspose
// Developer Intent: Generate an Excel file that contains a dedicated chart sheet displaying an OHLC stock chart based on data from another worksheet.
// Use Cases: Automated financial reports that keep raw data and OHLC visualizations on separate sheets. | Trading dashboards where each security’s OHLC chart is generated on its own chart sheet for easy distribution. | Batch processing of market data to produce printable OHLC charts without mixing them with source tables.
// AI Prompts: Write C# code with Aspose.Cells to add a chart sheet and create an OHLC stock chart from data on a different worksheet. | Show how to assign series and category ranges for an OHLC chart placed on a chart sheet using Aspose.Cells for .NET. | Explain how to customize the title, size, and position of an OHLC chart on a separate chart sheet with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, fills a data worksheet with Date, Open, High, Low, and Close values, adds a separate chart sheet, inserts an Open‑High‑Low‑Close (OHLC) stock chart, links the series to columns B‑E, sets the X‑axis to column A, adds a title, and saves the file as an XLSX workbook.
class OHLCChartSheetExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook (default worksheet at index 0 will hold the data)
            Workbook workbook = new Workbook();

            // Reference the default worksheet for data entry
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data for OHLC chart
            // Header row
            dataSheet.Cells["A1"].PutValue("Date");
            dataSheet.Cells["B1"].PutValue("Open");
            dataSheet.Cells["C1"].PutValue("High");
            dataSheet.Cells["D1"].PutValue("Low");
            dataSheet.Cells["E1"].PutValue("Close");

            // Sample rows (2‑6)
            dataSheet.Cells["A2"].PutValue("01/01/2023");
            dataSheet.Cells["B2"].PutValue(100);
            dataSheet.Cells["C2"].PutValue(110);
            dataSheet.Cells["D2"].PutValue(95);
            dataSheet.Cells["E2"].PutValue(105);

            dataSheet.Cells["A3"].PutValue("01/02/2023");
            dataSheet.Cells["B3"].PutValue(106);
            dataSheet.Cells["C3"].PutValue(115);
            dataSheet.Cells["D3"].PutValue(102);
            dataSheet.Cells["E3"].PutValue(112);

            dataSheet.Cells["A4"].PutValue("01/03/2023");
            dataSheet.Cells["B4"].PutValue(113);
            dataSheet.Cells["C4"].PutValue(118);
            dataSheet.Cells["D4"].PutValue(108);
            dataSheet.Cells["E4"].PutValue(110);

            dataSheet.Cells["A5"].PutValue("01/04/2023");
            dataSheet.Cells["B5"].PutValue(111);
            dataSheet.Cells["C5"].PutValue(119);
            dataSheet.Cells["D5"].PutValue(109);
            dataSheet.Cells["E5"].PutValue(117);

            dataSheet.Cells["A6"].PutValue("01/05/2023");
            dataSheet.Cells["B6"].PutValue(118);
            dataSheet.Cells["C6"].PutValue(125);
            dataSheet.Cells["D6"].PutValue(115);
            dataSheet.Cells["E6"].PutValue(122);

            // Add a new chart sheet to the workbook
            int chartSheetIndex = workbook.Worksheets.Add(SheetType.Chart);
            Worksheet chartSheet = workbook.Worksheets[chartSheetIndex];
            chartSheet.Name = "OHLC Chart";

            // Add an Open‑High‑Low‑Close (OHLC) stock chart to the chart sheet
            int chartIndex = chartSheet.Charts.Add(ChartType.StockOpenHighLowClose, 0, 0, 20, 10);
            Chart ohlcChart = chartSheet.Charts[chartIndex];

            // Set the data range for the series (Open, High, Low, Close)
            // Series data: columns B‑E, rows 2‑6 on the data sheet
            ohlcChart.NSeries.Add("Data!$B$2:$E$6", true);

            // Set the category (X‑axis) data – dates in column A
            ohlcChart.NSeries.CategoryData = "Data!$A$2:$A$6";

            // Optional: give the chart a title
            ohlcChart.Title.Text = "Sample OHLC Stock Chart";

            // Save the workbook with the chart sheet
            string outputPath = "OHLCChart.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
