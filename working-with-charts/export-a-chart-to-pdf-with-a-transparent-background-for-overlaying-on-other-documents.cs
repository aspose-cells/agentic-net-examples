// Title: Export a column chart to PDF with a transparent background using Aspose.Cells in C#
// AI Prompts: Write C# code that builds a column chart with Aspose.Cells, hides all chart and plot area borders, and saves only the chart as a PDF file that has no background color. | Adjust an existing Aspose.Cells workbook export routine so the generated PDF contains a fully transparent chart background suitable for overlay on other documents.
// Common Searches: how to create a transparent background for a chart exported to PDF with Aspose.Cells C# | Aspose.Cells C# export chart without background color for PDF overlay | remove chart area and plot area borders when saving Excel chart as PDF using Aspose.Cells | C# generate column chart and save as PDF with transparency using Aspose.Cells | export only chart to PDF with no background in Aspose.Cells library
// Tags: export chart to pdf aspnet cells | transparent chart background aspnet cells | hide chart borders aspnet cells | column chart pdf overlay aspnet cells | chart area transparency aspnet cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a column chart, disables chart and plot area borders, and saves the result as a PDF named ChartTransparent.pdf where the chart background is fully transparent, making it suitable for overlay on other documents.
class ExportChartToPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Optional: Remove chart borders for a cleaner overlay
            chart.ChartArea.Border.IsVisible = false;
            chart.PlotArea.Border.IsVisible = false;

            // Save the workbook as PDF
            string outputPath = "ChartTransparent.pdf";
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
