// Title: Export Excel Gantt Chart to PDF with Layout Preservation using Aspose.Cells for .NET
// Description: Load an Excel workbook that contains a Gantt chart, configure PdfSaveOptions.ExportDocumentStructure and RefreshChartCache to retain the chart's visual layout and data, and save the workbook as a PDF with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | Export Gantt chart to PDF | PdfSaveOptions | ExportDocumentStructure | RefreshChartCache | preserve chart layout | Excel to PDF conversion | chart rendering | project schedule PDF
// Common Searches: Aspose.Cells preserve chart layout PDF | Export Gantt chart Excel to PDF C# | PdfSaveOptions ExportDocumentStructure example | RefreshChartCache Aspose.Cells usage | Convert Excel workbook with charts to PDF .NET | Keep Gantt chart formatting when saving as PDF
// Developer Intent: Convert an Excel workbook that includes a Gantt chart into a PDF while keeping the chart’s visual layout and up‑to‑date data.
// Use Cases: Generate printable project schedule reports by converting Gantt‑chart Excel files to PDF. | Automate nightly creation of dashboard PDFs that contain multiple charts, ensuring visual fidelity. | Produce compliant financial or engineering documentation with embedded charts rendered exactly as in Excel. | Provide end‑users with offline PDF versions of interactive Excel timelines without losing formatting.
// AI Prompts: Write C# code using Aspose.Cells to export an Excel file with a Gantt chart to PDF, preserving layout via ExportDocumentStructure and RefreshChartCache. | Describe how ExportDocumentStructure and RefreshChartCache affect chart rendering when converting Excel to PDF with Aspose.Cells, and show a sample implementation. | Give troubleshooting steps for missing or distorted chart elements in a PDF exported from Excel using Aspose.Cells, even after setting ExportDocumentStructure.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsGanttExport
{
    // Load an Excel workbook that contains a Gantt chart, configure PdfSaveOptions.ExportDocumentStructure and RefreshChartCache to retain the chart's visual layout and data, and save the workbook as a PDF with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "GanttChart.xlsx";
                const string outputPath = "GanttChart.pdf";

                // Verify that the source workbook exists to avoid FileNotFoundException.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file '{inputPath}' was not found.");
                    return;
                }

                // Load the workbook that already contains a Gantt chart.
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options to preserve the document structure
                // (keeps layout of charts, shapes, etc.) and refresh the chart cache.
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    ExportDocumentStructure = true, // Preserve chart layout
                    RefreshChartCache = true        // Ensure chart data is up‑to‑date
                };

                // Save the entire workbook as a PDF file.
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook with Gantt chart exported to PDF successfully: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
