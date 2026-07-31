// Title: C# – Create a Separate Chart Sheet with an OHLC Stock Chart using Aspose.Cells
// Description: Demonstrates how to build a new workbook, populate a worksheet with date and OHLC values, add a dedicated chart sheet, and insert an Open‑High‑Low‑Close (OHLC) stock chart that references the data range. The example sets the category axis to dates, adds a chart title, and saves the file as OhlcChartSheet.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# OHLC chart | stock chart Aspose.Cells | chart sheet .NET | Open High Low Close chart | Excel OHLC example | Aspose.Cells chart sheet | financial chart C# | StockOpenHighLowClose | Aspose.Cells tutorial
// Common Searches: Aspose.Cells add OHLC chart sheet C# | How to create an Open‑High‑Low‑Close chart on a separate sheet using Aspose.Cells | C# example for chart sheet with OHLC stock chart | Create financial chart sheet Aspose.Cells .NET | Aspose.Cells chart sheet stock chart tutorial
// Developer Intent: Add a new chart sheet and insert an Open‑High‑Low‑Close (OHLC) stock chart that pulls data from another worksheet.
// Use Cases: Automate financial reports where each OHLC chart is placed on its own sheet for clear visualization. | Generate Excel templates that add separate OHLC chart sheets for multiple securities programmatically. | Export daily market data and provide analysts with a quick‑view chart sheet summarizing price movements.
// AI Prompts: Generate C# code with Aspose.Cells that creates a chart sheet and adds an OHLC chart from a worksheet named 'Data'. | Explain how to customize the OHLC chart title, axis labels, and styling after creating it on a chart sheet with Aspose.Cells. | Write a C# loop using Aspose.Cells to create separate OHLC chart sheets for a list of ticker symbols.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to build a new workbook, populate a worksheet with date and OHLC values, add a dedicated chart sheet, and insert an Open‑High‑Low‑Close (OHLC) stock chart that references the data range. The example sets the category axis to dates, adds a chart title, and saves the file as OhlcChartSheet.xlsx using Aspose.Cells for .NET.
class CreateOhlcChartSheet
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // -------------------------------------------------
        // Prepare sample OHLC data in the first worksheet
        // -------------------------------------------------
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Header row
        dataSheet.Cells["A1"].PutValue("Date");
        dataSheet.Cells["B1"].PutValue("Open");
        dataSheet.Cells["C1"].PutValue("High");
        dataSheet.Cells["D1"].PutValue("Low");
        dataSheet.Cells["E1"].PutValue("Close");

        // Sample rows
        dataSheet.Cells["A2"].PutValue("2023-01-01");
        dataSheet.Cells["B2"].PutValue(100);
        dataSheet.Cells["C2"].PutValue(110);
        dataSheet.Cells["D2"].PutValue(95);
        dataSheet.Cells["E2"].PutValue(105);

        dataSheet.Cells["A3"].PutValue("2023-01-02");
        dataSheet.Cells["B3"].PutValue(106);
        dataSheet.Cells["C3"].PutValue(115);
        dataSheet.Cells["D3"].PutValue(102);
        dataSheet.Cells["E3"].PutValue(112);

        dataSheet.Cells["A4"].PutValue("2023-01-03");
        dataSheet.Cells["B4"].PutValue(113);
        dataSheet.Cells["C4"].PutValue(118);
        dataSheet.Cells["D4"].PutValue(108);
        dataSheet.Cells["E4"].PutValue(110);

        dataSheet.Cells["A5"].PutValue("2023-01-04");
        dataSheet.Cells["B5"].PutValue(111);
        dataSheet.Cells["C5"].PutValue(119);
        dataSheet.Cells["D5"].PutValue(109);
        dataSheet.Cells["E5"].PutValue(117);

        dataSheet.Cells["A6"].PutValue("2023-01-05");
        dataSheet.Cells["B6"].PutValue(118);
        dataSheet.Cells["C6"].PutValue(125);
        dataSheet.Cells["D6"].PutValue(115);
        dataSheet.Cells["E6"].PutValue(122);

        // -------------------------------------------------
        // Insert a new chart sheet
        // -------------------------------------------------
        int chartSheetIndex = workbook.Worksheets.Add(SheetType.Chart);
        Worksheet chartSheet = workbook.Worksheets[chartSheetIndex];
        chartSheet.Name = "OHLC Chart";

        // -------------------------------------------------
        // Add an Open‑High‑Low‑Close (OHLC) stock chart
        // -------------------------------------------------
        // Position parameters are not relevant for a chart sheet,
        // but they must be supplied; using the full sheet area.
        int chartIndex = chartSheet.Charts.Add(ChartType.StockOpenHighLowClose, 0, 0, 30, 30);
        Chart chart = chartSheet.Charts[chartIndex];

        // Set the data range for the OHLC series (Open, High, Low, Close)
        chart.NSeries.Add("Data!B2:E6", true);
        // Set the category (X‑axis) data – the dates
        chart.NSeries.CategoryData = "Data!A2:A6";

        // Optional: give the chart a title
        chart.Title.Text = "Sample OHLC Stock Chart";

        // -------------------------------------------------
        // Save the workbook
        // -------------------------------------------------
        workbook.Save("OhlcChartSheet.xlsx");
    }
}
