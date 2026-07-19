// Title: Export a Workbook with an Updated Column Chart to PDF while Preserving Formatting – Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a column chart linked to sample data, configure PdfSaveOptions (RefreshChartCache and ExportDocumentStructure) and save the entire workbook as a PDF so the chart retains its original colors, layout, and accessibility information.
// Keywords: Aspose.Cells | .NET | C# | PDF export | chart formatting | RefreshChartCache | ExportDocumentStructure | column chart | save workbook as PDF | Excel to PDF conversion
// Common Searches: Aspose.Cells export chart to PDF preserving formatting | How to keep chart appearance when saving Excel as PDF in .NET | PdfSaveOptions RefreshChartCache example | Export workbook with chart to PDF using Aspose.Cells | C# save Excel chart to PDF with original style
// Developer Intent: Generate a PDF from a workbook that contains a column chart, ensuring the chart’s visual style and accessibility metadata are unchanged.
// Use Cases: Create printable reports that include up‑to‑date charts with exact colors and layout. | Automate dashboard exports from Excel to PDF for distribution to stakeholders. | Produce PDF invoices or financial statements where embedded charts must remain accessible and visually consistent.
// AI Prompts: Write C# code with Aspose.Cells to add a line chart, modify its data range, and export the workbook to PDF while preserving all chart styles. | Explain the impact of RefreshChartCache and ExportDocumentStructure on PDF output for charts in Aspose.Cells. | Show how to export only a specific worksheet that contains a chart to PDF, keeping the chart formatting intact.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Saving;

namespace AsposeCellsChartToPdf
{
    // Demonstrates how to create a workbook, add a column chart linked to sample data, configure PdfSaveOptions (RefreshChartCache and ExportDocumentStructure) and save the entire workbook as a PDF so the chart retains its original colors, layout, and accessibility information.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
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
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Refresh chart cache so the latest data is used
                RefreshChartCache = true,
                // Export document structure (optional, keeps accessibility info)
                ExportDocumentStructure = true
            };

            // Save the entire workbook (including the chart) as a PDF file
            workbook.Save("WorkbookWithChart.pdf", pdfOptions);

            Console.WriteLine("Workbook exported to PDF successfully.");
        }
    }
}
