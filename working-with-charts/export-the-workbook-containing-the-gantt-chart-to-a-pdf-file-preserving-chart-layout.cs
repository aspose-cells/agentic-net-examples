// Title: Export an Excel workbook with a Gantt chart to PDF while preserving chart layout – Aspose.Cells for .NET
// Description: Shows how to load a workbook that contains a Gantt chart, configure PdfSaveOptions (ExportDocumentStructure and RefreshChartCache) to retain the visual layout of charts and shapes, and save the workbook as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# PDF export | Gantt chart PDF | ExportDocumentStructure | RefreshChartCache | .NET Excel to PDF | preserve chart layout | chart rendering PDF | Aspose.Cells PdfSaveOptions | convert Excel with charts to PDF
// Common Searches: Aspose.Cells export Gantt chart to PDF | C# keep Excel chart layout when saving as PDF | PdfSaveOptions ExportDocumentStructure example | RefreshChartCache Aspose.Cells usage | how to preserve shapes in PDF export with Aspose.Cells
// Developer Intent: Generate a PDF file from an Excel workbook that includes a Gantt chart, ensuring the chart’s appearance matches the original worksheet.
// Use Cases: Produce printable project‑schedule reports where the Gantt chart must look identical to the Excel version. | Automate PDF snapshots of dashboards that contain multiple charts, shapes, and annotations. | Deliver client‑ready PDFs of Excel workbooks while maintaining exact positioning of all visual elements.
// AI Prompts: Write C# code with Aspose.Cells to load an Excel file containing a Gantt chart and save it as a PDF that keeps the chart layout intact. | Explain the impact of ExportDocumentStructure and RefreshChartCache on chart rendering during PDF conversion in Aspose.Cells. | Provide troubleshooting steps for distorted Gantt charts after exporting to PDF with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsGanttExport
{
    // Shows how to load a workbook that contains a Gantt chart, configure PdfSaveOptions (ExportDocumentStructure and RefreshChartCache) to retain the visual layout of charts and shapes, and save the workbook as a PDF using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            const string inputFile = "GanttChart.xlsx";
            const string outputFile = "GanttChart.pdf";

            // Verify that the source workbook exists before attempting to load it
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file '{inputFile}' was not found.");
                return;
            }

            try
            {
                // Load the workbook that already contains the Gantt chart
                Workbook workbook = new Workbook(inputFile);

                // Configure PDF save options to preserve the chart layout
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Export the document structure (charts, shapes, etc.) to keep layout intact
                    ExportDocumentStructure = true,
                    // Refresh chart cache to ensure the latest data is rendered
                    RefreshChartCache = true
                };

                // Save the entire workbook (including the Gantt chart) as a PDF file
                workbook.Save(outputFile, pdfOptions);

                Console.WriteLine($"Workbook with Gantt chart exported to PDF successfully: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
