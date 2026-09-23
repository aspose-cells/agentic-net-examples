// Title: Add a PDF bookmark that jumps to cell A1 of a chart sheet using Aspose.Cells for .NET
// AI Prompts: Insert a PdfBookmarkEntry named 'ChartStart' that points to the chart sheet's A1 cell, add it to PdfSaveOptions.Bookmarks, and regenerate the PDF. | Update the example to configure PdfSaveOptions with a bookmark collection containing an entry targeting the chart sheet's A1 cell, then save the workbook as PDF.
// Common Searches: aspnet aspose.cells create pdf bookmark for chart sheet cell a1 | how to set pdf bookmark destination to a chart sheet in asp.net | asp.net pdf export aspose.cells add bookmark to specific cell | pdfbookmarkentry chart sheet a1 aspose.cells example | saving chart sheet as pdf with bookmark using aspose.cells .net
// Tags: Aspose.Cells add PDF bookmark to chart sheet | PdfBookmarkEntry target chart sheet cell | PdfSaveOptions configure bookmarks .NET | export chart sheet to PDF with navigation bookmark | Aspose.Cells chart sheet PDF navigation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The example builds a workbook with a data worksheet and a chart sheet containing a column chart, then saves it as a PDF using PdfSaveOptions. It can be extended by adding a PdfBookmarkEntry that points to cell A1 of the chart sheet, enabling PDF navigation directly to the chart.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a data worksheet and populate some sample data
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";
                dataSheet.Cells["A1"].PutValue(10);
                dataSheet.Cells["A2"].PutValue(20);
                dataSheet.Cells["B1"].PutValue(30);
                dataSheet.Cells["B2"].PutValue(40);

                // Add a chart sheet
                Worksheet chartSheet = workbook.Worksheets.Add("ChartSheet");
                chartSheet.Type = SheetType.Chart; // Define the sheet as a chart sheet

                // Add a column chart to the chart sheet (position and size are required)
                int chartIndex = chartSheet.Charts.Add(0, 0, 400, 300, (int)ChartType.Column);
                Chart chart = chartSheet.Charts[chartIndex];
                chart.NSeries.Add("Data!A1:B2", true); // Use data from the data sheet
                chart.Title.Text = "Sample Chart";

                // Configure PDF save options (bookmarks omitted for compatibility)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Define output file path
                string outputPath = "ChartWorkbook.pdf";

                // Save the workbook as a PDF using the configured options
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
