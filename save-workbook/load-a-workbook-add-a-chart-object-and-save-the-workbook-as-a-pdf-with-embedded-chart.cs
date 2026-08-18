// Title: Add a Column Chart to an Excel Workbook and Export as PDF with Embedded Chart using Aspose.Cells for .NET
// Description: Loads an XLSX file, inserts sample data, creates a column chart on the first worksheet, sets the data range, enables RefreshChartCache in PdfSaveOptions, and saves the workbook as a PDF where the chart is rendered inside the document.
// Keywords: Aspose.Cells add chart C# | export Excel chart to PDF .NET | PdfSaveOptions RefreshChartCache | save workbook as PDF with chart | Aspose.Cells chart rendering PDF | C# Excel to PDF with embedded chart | Aspose.Cells tutorial US | Aspose.Cells Europe example
// Common Searches: Aspose.Cells embed chart in PDF C# | PdfSaveOptions RefreshChartCache example | Create column chart programmatically Aspose.Cells | Export Excel workbook to PDF with charts .NET | How to save Excel chart as PDF using Aspose
// Developer Intent: Generate a PDF from an Excel workbook that includes a newly created column chart.
// Use Cases: Produce sales summary PDFs that display product quantities with a column chart. | Automate financial dashboard exports to PDF while preserving all visual charts for client reports. | Create printable performance sheets where each page contains a chart generated from worksheet data.
// AI Prompts: Provide C# code to add a line chart to a worksheet and export the workbook as a PDF with the chart embedded using Aspose.Cells. | Explain how to configure PdfSaveOptions to refresh the chart cache so charts appear correctly in the exported PDF. | Show how to adjust chart size and position before saving the workbook to PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsChartPdfDemo
{
    // Loads an XLSX file, inserts sample data, creates a column chart on the first worksheet, sets the data range, enables RefreshChartCache in PdfSaveOptions, and saves the workbook as a PDF where the chart is rendered inside the document.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook from disk
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data for the chart (if not already present)
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(30);
            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["B3"].PutValue(45);
            worksheet.Cells["A4"].PutValue("Cherry");
            worksheet.Cells["B4"].PutValue(25);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Optional: refresh chart cache when saving to PDF
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.RefreshChartCache = true;

            // Save the entire workbook as a PDF; the chart will be embedded in the PDF
            string outputPath = "output.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine("Workbook saved as PDF with embedded chart at: " + outputPath);
        }
    }
}
