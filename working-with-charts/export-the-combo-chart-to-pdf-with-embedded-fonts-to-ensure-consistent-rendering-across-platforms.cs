// Title: Export a column‑line combo chart to PDF with embedded Windows fonts in C# using Aspose.Cells
// AI Prompts: Generate C# code that creates a workbook, adds a column‑line combo chart, and saves it as a PDF with standard Windows fonts embedded via Aspose.Cells. | Show how to set up Aspose.Cells PdfSaveOptions to embed fonts when exporting any chart to PDF in a .NET application. | Demonstrate adding a second data series as a line series to an existing column chart and exporting the combined chart to PDF with font embedding.
// Common Searches: Aspose.Cells C# export combo chart to PDF with embedded fonts | how to embed standard Windows fonts in PDF using Aspose.Cells PdfSaveOptions | create column and line combo chart in Aspose.Cells and save as PDF | PdfSaveOptions EmbedStandardWindowsFonts example C# | export chart to PDF with font embedding Aspose.Cells .NET
// Tags: combo chart PDF export Aspose.Cells | embed Windows fonts PdfSaveOptions | column line chart creation C# Aspose.Cells | font embedding in PDF export .NET | Aspose.Cells chart rendering to PDF

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The program builds a workbook with sample sales and profit data, creates a column‑line combo chart, configures PdfSaveOptions to embed standard Windows fonts, and saves the workbook as a PDF file named "ComboChart.pdf".
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the combo chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["C1"].PutValue("Profit");

            string[] months = { "Jan", "Feb", "Mar", "Apr" };
            double[] sales = { 12000, 15000, 18000, 21000 };
            double[] profit = { 3000, 3500, 4000, 4500 };

            for (int i = 0; i < months.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(months[i]);   // Column A
                sheet.Cells[i + 1, 1].PutValue(sales[i]);   // Column B
                sheet.Cells[i + 1, 2].PutValue(profit[i]);  // Column C
            }

            // Add a combo chart (Column + Line) to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // First series (Sales) as Column
            int salesSeriesIdx = chart.NSeries.Add("B2:B5", true);

            // Second series (Profit) as Line
            int profitSeriesIdx = chart.NSeries.Add("C2:C5", true);
            chart.NSeries[profitSeriesIdx].Type = ChartType.Line;

            // Set category (X‑axis) data
            chart.NSeries.CategoryData = "A2:A5";

            // Optional: set chart title
            chart.Title.Text = "Monthly Sales and Profit";

            // Configure PDF save options to embed standard Windows fonts
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                EmbedStandardWindowsFonts = true
            };

            // Save the workbook (including the combo chart) as PDF
            workbook.Save("ComboChart.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
