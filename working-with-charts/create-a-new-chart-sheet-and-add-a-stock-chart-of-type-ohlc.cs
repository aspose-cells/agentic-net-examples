// Title: How to add a separate chart sheet with an OHLC (Open‑High‑Low‑Close) stock chart using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a new workbook, adds a dedicated chart sheet, inserts an OHLC stock chart, links it to data on another worksheet, and saves the file. | Show how to fill Open, High, Low, and Close columns in a worksheet and reference that range for an OHLC chart placed on a separate chart sheet with Aspose.Cells. | Provide an example that sets the OHLC chart’s position and dimensions in pixels and ensures the output directory exists before saving the workbook.
// Common Searches: Aspose.Cells create chart sheet with OHLC stock chart in C# | C# generate Open High Low Close chart on separate worksheet using Aspose.Cells | How to bind an OHLC chart to a data range in Aspose.Cells .NET | Set chart size in pixels Aspose.Cells C# example | Save workbook to specific folder after creating OHLC chart with Aspose.Cells
// Tags: chart sheet creation OHLC Aspose.Cells | populate OHLC data worksheet C# | bind chart to range Sheet1!A2:D6 Aspose.Cells | chart pixel dimensions Aspose.Cells | save workbook OHLCChart.xlsx Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The program creates a new workbook, adds a dedicated chart sheet named "OHLCChart", inserts an Open‑High‑Low‑Close stock chart at pixel coordinates (0,0) with size 400×500, fills sample Open, High, Low, and Close values in the first worksheet, binds the chart to the range A2:D6, ensures the output directory exists, and saves the workbook as OHLCChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook with a default worksheet
            Workbook workbook = new Workbook();

            // Add a worksheet that will host the OHLC chart
            Worksheet chartSheet = workbook.Worksheets.Add("OHLCChart");

            // Add an OHLC chart (Open‑High‑Low‑Close) to the worksheet
            // Position (row, column) and size (height, width) are specified in pixels
            int chartIndex = chartSheet.Charts.Add(0, 0, 400, 500, (int)ChartType.StockOpenHighLowClose);
            Chart ohlcChart = chartSheet.Charts[chartIndex];

            // Fill sample data into the first worksheet (Sheet1)
            worksheetData(workbook.Worksheets[0]);

            // Set the data range for the chart (A2:D6)
            ohlcChart.NSeries.Add("Sheet1!A2:D6", true);

            // Save the workbook to a file
            string outputPath = "OHLCChart.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to fill sample data into the first worksheet
    static void worksheetData(Worksheet sheet)
    {
        // Headers
        sheet.Cells["A1"].PutValue("Open");
        sheet.Cells["B1"].PutValue("High");
        sheet.Cells["C1"].PutValue("Low");
        sheet.Cells["D1"].PutValue("Close");

        // Sample numeric data
        double[,] data = new double[,]
        {
            { 100, 110, 95, 105 },
            { 105, 115, 100, 110 },
            { 110, 120, 105, 115 },
            { 115, 125, 110, 120 },
            { 120, 130, 115, 125 }
        };

        for (int i = 0; i < data.GetLength(0); i++)
        {
            for (int j = 0; j < data.GetLength(1); j++)
            {
                // Rows are zero‑based; add 1 to skip header row
                sheet.Cells[i + 1, j].PutValue(data[i, j]);
            }
        }
    }
}
