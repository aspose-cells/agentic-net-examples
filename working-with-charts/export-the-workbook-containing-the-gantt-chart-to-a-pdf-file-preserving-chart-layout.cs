// Title: Export an Excel workbook with a Gantt chart to PDF while preserving chart layout using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a .xlsx file containing a Gantt chart and saves it as a PDF with the original layout using Aspose.Cells. | Show how to configure PdfSaveOptions in Aspose.Cells to disable one‑page‑per‑sheet scaling and keep charts and shapes intact during PDF conversion. | Provide a complete example that converts an Excel workbook to PDF without altering the positioning of Gantt chart elements.
// Common Searches: Aspose.Cells C# export Excel Gantt chart to PDF without losing layout | How to keep chart positions when saving Excel as PDF with Aspose.Cells | PdfSaveOptions OnePagePerSheet false example for preserving sheet design | Convert Excel workbook with Gantt diagram to PDF preserving shapes Aspose .NET
// Tags: Aspose.Cells preserve chart layout PDF | PdfSaveOptions disable one page per sheet | C# convert Gantt chart workbook to PDF | Excel workbook PDF conversion keeping shapes | Aspose.Cells PDF conversion settings .NET

using Aspose.Cells;
using System;

// Loads an Excel workbook that contains a Gantt chart and saves it as a PDF using Aspose.Cells, configuring PdfSaveOptions (OnePagePerSheet = false, AllColumnsInOnePagePerSheet = false) to retain the original chart layout and shapes.
class Program
{
    static void Main()
    {
        // Load the workbook that already contains the Gantt chart
        Workbook workbook = new Workbook("GanttChart.xlsx");

        // Configure PDF save options to keep the original layout (charts, shapes, etc.)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Preserve the sheet layout as it appears in Excel
            OnePagePerSheet = false,
            AllColumnsInOnePagePerSheet = false
        };

        // Export the workbook to PDF while preserving the Gantt chart layout
        workbook.Save("GanttChart.pdf", pdfOptions);
    }
}
