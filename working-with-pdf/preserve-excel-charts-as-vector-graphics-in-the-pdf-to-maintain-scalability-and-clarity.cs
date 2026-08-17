// Title: Export Excel Chart to PDF as Vector Graphics with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a column chart, and use Aspose.Cells Chart.ToPdf method to render the chart as vector graphics, ensuring the PDF remains scalable and crisp at any zoom level.
// Keywords: Aspose.Cells | C# chart to PDF | vector PDF export | Chart.ToPdf | preserve chart quality | Excel chart PDF vector | Aspose.Cells .NET | scalable PDF graphics | export column chart PDF | high‑resolution PDF chart
// Common Searches: Aspose.Cells export chart as vector PDF | Chart.ToPdf vector output C# | How to keep Excel chart sharp in PDF | Convert Excel chart to scalable PDF .NET | Save Excel chart as vector graphics PDF | C# export chart to PDF without rasterizing
// Developer Intent: Generate a PDF file from an Excel chart where the chart is rendered as vector graphics for loss‑less scaling.
// Use Cases: Embedding high‑resolution charts in client‑facing PDF reports | Creating printable financial dashboards with crisp graphics | Automating batch conversion of workbook charts to individual vector PDFs for publishing | Generating PDFs for regulatory filings where chart clarity is mandatory
// AI Prompts: Write C# code using Aspose.Cells to export a pie chart to a vector PDF with custom page size. | Show how to configure PdfSaveOptions to preserve vector rendering when saving an entire workbook to PDF. | Explain how to loop through all charts in a workbook and save each as a separate vector PDF file using Aspose.Cells. | Provide an example of adding a legend and data labels to a chart before exporting it as vector PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartToPdfVector
{
    // Demonstrates how to create a workbook, add a column chart, and use Aspose.Cells Chart.ToPdf method to render the chart as vector graphics, ensuring the PDF remains scalable and crisp at any zoom level.
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
            worksheet.Cells["A2"].PutValue("Fruits");
            worksheet.Cells["A3"].PutValue("Vegetables");
            worksheet.Cells["A4"].PutValue("Grains");

            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(50);
            worksheet.Cells["B3"].PutValue(30);
            worksheet.Cells["B4"].PutValue(20);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Export the chart to PDF.
            // The ToPdf method renders the chart as vector graphics,
            // preserving scalability and clarity in the resulting PDF.
            chart.ToPdf("ChartVectorOutput.pdf");

            Console.WriteLine("Chart exported to PDF as vector graphics successfully.");
        }
    }
}
