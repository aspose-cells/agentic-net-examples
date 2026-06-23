using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells;

namespace AsposeCellsChartValidation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Validate that the worksheet contains at least one chart before exporting
            if (worksheet.Charts.Count > 0)
            {
                // Export the entire workbook to PDF
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Ensure chart data is refreshed during export
                    RefreshChartCache = true
                };

                workbook.Save("WorksheetWithChart.pdf", pdfOptions);
                Console.WriteLine("Workbook exported to PDF successfully.");
            }
            else
            {
                Console.WriteLine("No charts found in the worksheet. PDF export aborted.");
            }
        }
    }
}