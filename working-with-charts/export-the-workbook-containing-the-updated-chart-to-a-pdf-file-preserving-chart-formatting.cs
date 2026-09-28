// Title: Export a workbook with a formatted column chart to PDF while preserving chart layout using Aspose.Cells in C#
// AI Prompts: Write C# code that creates a new workbook, inserts a column chart with a title, sets PdfSaveOptions (RefreshChartCache = true and ExportDocumentStructure = true), and saves the workbook as a PDF file. | Show how to use Aspose.Cells PdfSaveOptions to retain chart formatting when converting an Excel workbook containing charts to PDF in C#.
// Common Searches: Aspose.Cells C# export Excel workbook with chart to PDF preserving chart formatting | How to keep column chart layout when saving workbook as PDF using Aspose.Cells | PdfSaveOptions RefreshChartCache true example in C# | ExportDocumentStructure option for PDF conversion with charts Aspose.Cells | C# code to save workbook with chart to PDF using Aspose.Cells
// Tags: Aspose.Cells PDF export with chart formatting | PdfSaveOptions RefreshChartCache C# | ExportDocumentStructure option Aspose.Cells | column chart PDF conversion Aspose.Cells | C# workbook to PDF preserving charts

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Saving;

namespace AsposeCellsExportWorkbookToPdf
{
    // The program creates a new workbook, adds sample data, inserts a column chart with a title, configures PdfSaveOptions to refresh the chart cache and retain document structure, and saves the workbook as 'WorkbookWithChart.pdf', preserving the chart's appearance in the PDF.
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
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(50);
            worksheet.Cells["B3"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Optional: Apply some formatting to the chart (e.g., title)
            chart.Title.Text = "Sample Chart";

            // Create PDF save options to preserve chart formatting
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
            {
                // Refresh chart cache ensures the latest data and formatting are used
                RefreshChartCache = true,
                // ExportDocumentStructure retains the structure of the document
                ExportDocumentStructure = true
            };

            // Save the entire workbook (including the chart) to a PDF file
            workbook.Save("WorkbookWithChart.pdf", pdfSaveOptions);

            Console.WriteLine("Workbook exported to PDF successfully.");
        }
    }
}
