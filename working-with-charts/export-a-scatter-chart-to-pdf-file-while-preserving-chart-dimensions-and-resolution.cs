// Title: Export a Scatter Chart to PDF with Exact Size and 300 DPI Using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export scatter chart to PDF preserving chart size | set chart object width and height before PDF conversion Aspose.Cells | configure 300 DPI image resolution for PDF output in Aspose.Cells | keep Excel chart layout unchanged when saving as PDF with Aspose.Cells
// Tags: scatter chart PDF export Aspose.Cells .NET | chart object dimensions Aspose.Cells | default image resolution 300 DPI Aspose.Cells | OnePagePerSheet PDF option Aspose.Cells | excel to pdf rasterization Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace ScatterChartToPdfDemo
{
    // The example creates a workbook, fills X‑Y data, adds a scatter chart, sets the chart object's left, top, width, and height in points, configures the workbook's default image resolution to 300 dpi, enables OnePagePerSheet in PdfSaveOptions, and saves both an .xlsx and a .pdf file where the chart retains its exact dimensions and high‑resolution rendering.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for the generated files (relative to the executable folder)
            const string excelPath = "ScatterChart.xlsx";
            const string pdfPath = "ScatterChart.pdf";

            try
            {
                // ------------------------------------------------------------
                // 1. Create a new workbook and populate it with sample data
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Header
                cells["A1"].PutValue("X");
                cells["B1"].PutValue("Y");

                // Sample data points (10 points)
                double[] xValues = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
                double[] yValues = { 2, 4, 5, 7, 11, 13, 14, 18, 20, 22 };

                for (int i = 0; i < xValues.Length; i++)
                {
                    cells[i + 1, 0].PutValue(xValues[i]); // Column A
                    cells[i + 1, 1].PutValue(yValues[i]); // Column B
                }

                // ------------------------------------------------------------
                // 2. Add a Scatter chart and bind it to the data range
                // ------------------------------------------------------------
                // Add the chart: (type, upper-left row, upper-left column, lower-right row, lower-right column)
                int chartIndex = sheet.Charts.Add(ChartType.Scatter, 12, 0, 30, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the series
                chart.NSeries.Add("B2:B11", true);
                chart.NSeries.CategoryData = "A2:A11";

                // Give the chart a title
                chart.Title.Text = "Sample Scatter Chart";

                // ------------------------------------------------------------
                // 3. Preserve the exact chart dimensions (in points)
                // ------------------------------------------------------------
                chart.ChartObject.Left = 50;    // 50 points from the left edge of the sheet
                chart.ChartObject.Top = 20;     // 20 points from the top edge of the sheet
                chart.ChartObject.Width = 500;  // Desired width in points
                chart.ChartObject.Height = 350; // Desired height in points

                // ------------------------------------------------------------
                // 4. Save the workbook as an Excel file (optional, for verification)
                // ------------------------------------------------------------
                workbook.Save(excelPath, SaveFormat.Xlsx);
                Console.WriteLine($"Excel file saved to '{excelPath}'.");

                // ------------------------------------------------------------
                // 5. Configure PDF save options and set image resolution
                // ------------------------------------------------------------
                // Set default image resolution for the workbook (applies to PDF rasterization)
                workbook.Settings.DefaultImageResolution = 300;

                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Render each worksheet on a single PDF page (keeps layout intact)
                    OnePagePerSheet = true
                };

                // ------------------------------------------------------------
                // 6. Export the workbook (including the scatter chart) to PDF
                // ------------------------------------------------------------
                workbook.Save(pdfPath, pdfOptions);
                Console.WriteLine($"PDF file with preserved chart dimensions saved to '{pdfPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the scatter chart PDF:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
