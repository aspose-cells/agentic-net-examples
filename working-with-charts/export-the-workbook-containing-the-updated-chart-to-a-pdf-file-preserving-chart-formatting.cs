// Title: Export Aspose.Cells workbook with a column chart to PDF while preserving formatting (C#)
// Description: Creates a workbook, adds sample data and a column chart, configures PdfSaveOptions (RefreshChartCache and ExportDocumentStructure) to keep the chart up‑to‑date and its visual layout, then saves the whole workbook as a PDF.
// Keywords: Aspose.Cells | C# PDF export | chart to PDF | RefreshChartCache | ExportDocumentStructure | .NET Excel to PDF | preserve chart appearance | column chart export
// Common Searches: Aspose.Cells export chart to PDF C# | how to keep chart formatting when saving Excel as PDF | PdfSaveOptions RefreshChartCache example | ExportDocumentStructure effect on PDF accessibility | save workbook with chart as PDF using Aspose.Cells
// Developer Intent: Generate a PDF file from a workbook that contains a column chart, ensuring the chart’s design and data are retained.
// Use Cases: Produce printable PDF reports that include sales or KPI charts with the same look as the Excel source. | Automate dashboard distribution where charts must reflect the latest data at export time. | Create accessible PDFs by preserving document structure for screen‑reader compatibility.
// AI Prompts: Show C# code to export a workbook with multiple charts to a single PDF while maintaining each chart’s formatting. | Explain how RefreshChartCache influences chart rendering in PDF output with Aspose.Cells. | Give examples of using ExportDocumentStructure to improve PDF accessibility for Excel charts.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Saving; // Namespace for PdfSaveOptions

// Creates a workbook, adds sample data and a column chart, configures PdfSaveOptions (RefreshChartCache and ExportDocumentStructure) to keep the chart up‑to‑date and its visual layout, then saves the whole workbook as a PDF.
class ExportWorkbookChartToPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Fruits");
        sheet.Cells["A3"].PutValue("Vegetables");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["B3"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Configure PDF save options to preserve chart formatting
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.RefreshChartCache = true;          // Ensure chart data is up‑to‑date
        pdfOptions.ExportDocumentStructure = true;    // Preserve document structure (optional)

        // Save the entire workbook, including the chart, to a PDF file
        workbook.Save("WorkbookWithChart.pdf", pdfOptions);

        Console.WriteLine("Workbook exported to PDF successfully.");
    }
}
