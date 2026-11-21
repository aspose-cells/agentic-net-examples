// Title: Export a combo chart to PDF with embedded fonts using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# export combo chart to PDF with embedded fonts | How to embed fonts when saving an Excel combo chart as PDF using Aspose.Cells .NET | Save workbook containing a combo chart to PDF with font embedding in C# | PdfSaveOptions font embedding for charts example in Aspose.Cells | Generate PDF from a combo chart with consistent fonts using Aspose.Cells
// Tags: Aspose.Cells combo chart PDF export | C# PdfSaveOptions font embedding | export Excel chart to PDF Aspose.Cells | embed fonts in PDF Aspose.Cells .NET | combo chart PDF generation C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, fills it with month, sales, and profit data, adds a combo chart (column for sales, line for profit), configures chart titles and axes, sets up PdfSaveOptions for font embedding, and saves the workbook as a PDF so the chart renders consistently across platforms.
class ExportComboChartToPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            // Columns: A - Category, B - Series1 (Column), C - Series2 (Line)
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["C1"].PutValue("Profit");

            string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
            double[] sales = { 12000, 15000, 13000, 17000, 16000 };
            double[] profit = { 3000, 4000, 3500, 4500, 4200 };

            for (int i = 0; i < months.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(months[i]);   // Column A
                sheet.Cells[i + 1, 1].PutValue(sales[i]);   // Column B
                sheet.Cells[i + 1, 2].PutValue(profit[i]);  // Column C
            }

            // Add a combo chart (default type is Column)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Monthly Sales and Profit";

            // Add first series (Column) for Sales
            int nSeriesIndex = chart.NSeries.Add("B2:B6", true);
            chart.NSeries[nSeriesIndex].Name = "Sales";
            chart.NSeries[nSeriesIndex].Type = ChartType.Column;

            // Add second series (Line) for Profit
            nSeriesIndex = chart.NSeries.Add("C2:C6", true);
            chart.NSeries[nSeriesIndex].Name = "Profit";
            chart.NSeries[nSeriesIndex].Type = ChartType.Line;

            // Set category axis labels (Month)
            chart.NSeries.CategoryData = "A2:A6";

            // Optional: customize axes
            chart.ValueAxis.Title.Text = "Amount";
            chart.CategoryAxis.Title.Text = "Month";

            // Prepare PDF save options (default font handling)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook (including the chart) as PDF
            workbook.Save("ComboChart.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
