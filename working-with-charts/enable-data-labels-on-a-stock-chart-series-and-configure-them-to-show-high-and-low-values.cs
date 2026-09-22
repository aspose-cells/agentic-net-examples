// Title: How to enable data labels showing high and low values on a Stock Open‑High‑Low‑Close chart with Aspose.Cells for .NET (C#)
// AI Prompts: Generate a StockOpenHighLowClose chart from worksheet data and turn on data labels that display the high and low values for each point using Aspose.Cells in C#. | Configure the Series.DataLabels.ShowValue property to true for a stock chart series so the Excel file created with Aspose.Cells shows value labels.
// Common Searches: Aspose.Cells C# add data labels to stock chart high low values | show high and low values on StockOpenHighLowClose chart using Aspose.Cells | enable value labels for OHLC chart in .NET Excel library | C# Aspose.Cells series data labels ShowValue example | how to display high low labels on Excel stock chart programmatically
// Tags: Aspose.Cells StockOpenHighLowClose data labels | Series.DataLabels.ShowValue C# | Excel stock chart value labels Aspose.Cells | C# generate OHLC chart with Aspose.Cells | enable high low labels in Excel chart .NET

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a workbook, fills it with sample stock data, adds a StockOpenHighLowClose chart referencing the Open, High, Low, and Close columns, enables data labels by setting Series.DataLabels.ShowValue to true, and saves the workbook as an Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and name it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Fill header row
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Open");
            sheet.Cells["C1"].PutValue("High");
            sheet.Cells["D1"].PutValue("Low");
            sheet.Cells["E1"].PutValue("Close");

            // Populate sample stock data
            DateTime startDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(startDate.AddDays(i)); // Date
                sheet.Cells[i + 1, 1].PutValue(100 + i * 2);          // Open
                sheet.Cells[i + 1, 2].PutValue(105 + i * 2);          // High
                sheet.Cells[i + 1, 3].PutValue(95 + i * 2);           // Low
                sheet.Cells[i + 1, 4].PutValue(102 + i * 2);          // Close
            }

            // Add a Stock chart (Open-High-Low-Close) to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.StockOpenHighLowClose, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Sample Stock Chart";

            // Add a series that uses Open, High, Low, Close columns
            int seriesIndex = chart.NSeries.Add("Data!$B$2:$E$6", true);
            Series series = chart.NSeries[seriesIndex];

            // Enable data labels (show values)
            // Note: In recent Aspose.Cells versions, setting ShowValue is sufficient.
            series.DataLabels.ShowValue = true; // Shows the data values (high/low for stock chart)

            // Define output file path
            string outputPath = "StockChartWithDataLabels.xlsx";

            // Ensure the directory exists before saving (if a directory is specified)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
