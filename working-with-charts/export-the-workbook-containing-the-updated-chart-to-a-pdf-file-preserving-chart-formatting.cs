// Title: Export Workbook with Column Chart to PDF – Preserve Chart Formatting (Aspose.Cells for .NET)
// Description: Demonstrates how to create a workbook, add a column chart, configure PdfSaveOptions (RefreshChartCache, ExportDocumentStructure) and save the entire workbook as a PDF while keeping the chart’s layout and styling intact.
// Keywords: Aspose.Cells export chart to PDF | preserve chart formatting PDF | PdfSaveOptions RefreshChartCache | ExportDocumentStructure Aspose.Cells | C# save workbook as PDF with chart | Aspose.Cells chart PDF export .NET
// Common Searches: Aspose.Cells export Excel chart to PDF | keep chart appearance when saving PDF with Aspose.Cells | RefreshChartCache effect on PDF output | ExportDocumentStructure for accessible PDF Aspose.Cells | C# code to save workbook with chart as PDF
// Developer Intent: Save an Excel workbook that contains a column chart as a PDF file without losing the chart’s visual formatting.
// Use Cases: Automated generation of PDF reports that include styled charts for business dashboards. | Creating printable invoices or statements where sales data is visualized in a chart that must match the Excel source. | Building compliance‑ready PDFs with accessible document structure while retaining chart design.
// AI Prompts: Generate C# code using Aspose.Cells to add a line chart and export the workbook to PDF while preserving chart formatting. | Explain the role of RefreshChartCache and ExportDocumentStructure in PdfSaveOptions for chart‑rich PDF exports. | Provide a step‑by‑step guide to customize chart colors, titles, and styles before saving the workbook as a PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Saving;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a column chart, configure PdfSaveOptions (RefreshChartCache, ExportDocumentStructure) and save the entire workbook as a PDF while keeping the chart’s layout and styling intact.
    public class ExportWorkbookWithChartToPdf
    {
        public static void Run()
        {
            try
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

                // Optional: customize chart appearance (title, style, etc.)
                chart.Title.Text = "Sample Chart";
                chart.Title.Font.Size = 12;
                chart.Title.Font.IsBold = true;

                // Prepare PDF save options to ensure chart formatting is retained
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Refresh chart cache so that the latest data and formatting are used
                    RefreshChartCache = true,

                    // Preserve document structure (optional, but keeps accessibility features)
                    ExportDocumentStructure = true
                };

                // Save the entire workbook (including the chart) as a PDF file
                string outputPath = "WorkbookWithChart.pdf";
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook exported to PDF successfully: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during PDF export: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
