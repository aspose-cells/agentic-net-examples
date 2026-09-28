// Title: Create a column chart with two data series and freeze the source rows using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that adds two separate series to a column chart, each referencing its own X‑value range, with Aspose.Cells. | Show how to freeze the first five rows of a worksheet after inserting chart data using the FreezePanes method in Aspose.Cells for .NET. | Provide a complete Aspose.Cells C# example that builds a workbook, fills series data, creates a column chart, applies row freezing, and saves the file.
// Common Searches: how to add multiple series to a column chart in Aspose.Cells C# | Aspose.Cells FreezePanes to lock header rows for chart data | C# Aspose.Cells example creating column chart with separate X values | freeze rows 1-5 in worksheet after chart creation Aspose.Cells | populate chart data ranges and freeze rows using Aspose.Cells .NET
// Tags: add multiple series to column chart Aspose.Cells | set XValues for chart series Aspose.Cells C# | freeze worksheet rows using FreezePanes Aspose.Cells | populate chart data ranges Aspose.Cells .NET | save workbook as Excel file Aspose.Cells C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// // Demonstrates creating a workbook, populating two data series, adding a column chart with distinct X‑value ranges, freezing the first five rows, and saving the file as ChartWithMultipleSeries.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Populate sample data for two series
            // -------------------------------------------------
            // Header row
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["C1"].PutValue("Category");
            sheet.Cells["D1"].PutValue("Series 2");

            // Data rows (rows 2‑5)
            for (int i = 0; i < 4; i++)
            {
                // Category labels (same for both series)
                sheet.Cells[i + 1, 0].PutValue("Item " + (i + 1));

                // Series 1 values
                sheet.Cells[i + 1, 1].PutValue(i + 2); // 2,3,4,5

                // Series 2 values
                sheet.Cells[i + 1, 3].PutValue((i + 1) * 3); // 3,6,9,12
            }

            // -------------------------------------------------
            // Add a column chart to the worksheet
            // -------------------------------------------------
            // Parameters: chart type, upper‑left row/col, lower‑right row/col
            int chartIdx = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // -------------------------------------------------
            // Add first series (Series 1)
            // -------------------------------------------------
            int seriesIdx1 = chart.NSeries.Add("B2:B5", true);
            chart.NSeries[seriesIdx1].Name = "Series 1";
            chart.NSeries[seriesIdx1].XValues = "A2:A5";

            // -------------------------------------------------
            // Add second series (Series 2)
            // -------------------------------------------------
            int seriesIdx2 = chart.NSeries.Add("D2:D5", true);
            chart.NSeries[seriesIdx2].Name = "Series 2";
            chart.NSeries[seriesIdx2].XValues = "C2:C5";

            // -------------------------------------------------
            // Freeze rows that contain the series data (rows 1‑5)
            // -------------------------------------------------
            // FreezePanes(row, column, freezableRows, freezableColumns)
            // Row index is zero‑based; to freeze rows 0‑4 set row = 5.
            sheet.FreezePanes(5, 0, 5, 0);

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("ChartWithMultipleSeries.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
