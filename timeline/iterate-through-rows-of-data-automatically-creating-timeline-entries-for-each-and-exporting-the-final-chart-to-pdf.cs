// Title: Create a timeline line chart from Excel rows and export it to PDF with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells generate timeline chart from Excel data and save as PDF | How to add a line chart with dynamic range to a worksheet using Aspose.Cells | Export Aspose.Cells workbook containing a timeline chart to PDF in .NET | Iterate Excel rows to set chart series X and Y values with Aspose.Cells C# | Create sample Excel file and produce timeline PDF chart using Aspose.Cells
// Tags: timeline line chart generation Aspose.Cells | dynamic chart series range Aspose.Cells C# | export workbook to PDF Aspose.Cells | sample Excel workbook creation Aspose.Cells | chart plot area formatting Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example checks for a source Excel file, creates a sample workbook if missing, loads it, determines the last used row, adds a line chart that serves as a timeline, assigns dates from column A as X‑values and numeric data from column B as Y‑values, applies a light‑yellow background and gray border to the plot area, and finally saves the workbook with the chart as a PDF file.
class TimelineChartGenerator
{
    static void Main()
    {
        try
        {
            const string dataFile = "Data.xlsx";
            const string outputFile = "TimelineChart.pdf";

            // Ensure the source workbook exists; create a simple one if it does not.
            if (!File.Exists(dataFile))
            {
                CreateSampleWorkbook(dataFile);
            }

            // Load the workbook that contains the source data
            Workbook workbook = new Workbook(dataFile);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the last used row (assuming the first row contains headers)
            int lastRow = sheet.Cells.MaxDataRow + 1; // MaxDataRow is zero‑based, convert to 1‑based

            // Add a line chart that will serve as the timeline
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Timeline Chart";

            // Define the range for X‑axis (dates in column A) and Y‑axis (values in column B)
            string xRange = $"{sheet.Name}!$A$2:$A${lastRow}";
            string yRange = $"{sheet.Name}!$B$2:$B${lastRow}";

            // Add a single series that uses the defined X and Y ranges
            chart.NSeries.Add(yRange, true);
            chart.NSeries[0].XValues = xRange;

            // Optional: format the chart appearance
            chart.PlotArea.Area.ForegroundColor = Color.LightYellow;
            chart.PlotArea.Border.IsVisible = true;
            chart.PlotArea.Border.Color = Color.Gray;

            // Save the workbook (including the chart) as a PDF file
            workbook.Save(outputFile, SaveFormat.Pdf);
            Console.WriteLine($"Chart successfully saved to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Creates a simple workbook with sample date and value data
    private static void CreateSampleWorkbook(string filePath)
    {
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Header
        ws.Cells["A1"].PutValue("Date");
        ws.Cells["B1"].PutValue("Value");

        // Sample data
        DateTime start = DateTime.Today.AddDays(-5);
        for (int i = 0; i < 6; i++)
        {
            ws.Cells[i + 1, 0].PutValue(start.AddDays(i));
            ws.Cells[i + 1, 1].PutValue(i * 10 + 5);
        }

        wb.Save(filePath);
    }
}
