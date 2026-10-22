// Title: Create a timeline line chart from worksheet rows and export the workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads date and value rows from a worksheet, builds a line chart representing a timeline, and saves the workbook as a PDF with Aspose.Cells. | Adapt the sample to produce a scatter chart with a date‑axis timeline and export the result to PDF using Aspose.Cells. | Enhance the timeline chart by applying custom date‑axis formatting and legend settings before saving the workbook to PDF.
// Common Searches: how to build a timeline line chart from Excel rows and export to PDF with Aspose.Cells C# | Aspose.Cells .NET export chart with dynamic data range to PDF | create date‑axis timeline chart in Aspose.Cells and save as PDF file
// Tags: create line chart Aspose.Cells | export chart to PDF Aspose.Cells | dynamic data range chart Aspose.Cells | populate worksheet rows Aspose.Cells | timeline chart using date axis Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a new workbook, fills column A with sequential dates and column B with numeric values, adds a line chart that uses these ranges as a timeline, optionally customizes the chart, and saves the workbook—including the chart—as a PDF file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add header row
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");

            // Populate sample data (replace with your own data source as needed)
            DateTime startDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(startDate.AddDays(i)); // Date column
                sheet.Cells[i + 1, 1].PutValue(i * 10);               // Value column
            }

            // Determine the last row with data
            int lastDataRow = sheet.Cells.MaxDataRow;

            // Add a Line chart to the worksheet (used as a timeline representation)
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
            Chart timelineChart = sheet.Charts[chartIndex];

            // Define the data ranges for dates and values
            string dateRange = $"A2:A{lastDataRow + 1}";
            string valueRange = $"B2:B{lastDataRow + 1}";

            // Add the series to the chart (date range first, then value range)
            timelineChart.NSeries.Add($"{dateRange},{valueRange}", true);
            timelineChart.NSeries[0].Name = "Sample Timeline";

            // Optional: set chart title
            timelineChart.Title.Text = "Timeline Chart";

            // Export the workbook (including the chart) to PDF
            string outputPath = "TimelineChart.pdf";
            workbook.Save(outputPath, SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
