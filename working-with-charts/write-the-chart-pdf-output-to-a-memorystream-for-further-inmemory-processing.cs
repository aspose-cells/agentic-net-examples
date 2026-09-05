// Title: Write an Aspose.Cells chart to a PDF MemoryStream for in‑memory processing in C#
// AI Prompts: Generate C# code that creates a worksheet, adds a column chart, and exports the chart to a PDF directly into a MemoryStream using Aspose.Cells. | Show how to retrieve the PDF byte array from a chart exported with chart.ToPdf without writing a file to disk. | Demonstrate resetting the MemoryStream position after exporting a chart to PDF for subsequent reading.
// Common Searches: Aspose.Cells export chart to PDF memory stream C# example | How to get PDF bytes from an Excel chart using Aspose.Cells | Create PDF from chart in memory with Aspose.Cells without saving file | C# Aspose.Cells chart.ToPdf to MemoryStream usage guide
// Tags: Aspose.Cells chart.ToPdf memory stream | in‑memory PDF generation from Excel chart | retrieve PDF byte array Aspose.Cells chart | C# export chart to PDF without file | chart PDF output using Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills it with sample data, adds a column chart, and uses chart.ToPdf to write the chart as a PDF into a MemoryStream. It then resets the stream position and extracts the PDF bytes for further in‑memory processing, avoiding any file system writes.
class ChartPdfToMemoryStreamDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
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

        // Create a memory stream to hold the PDF output
        using (MemoryStream pdfStream = new MemoryStream())
        {
            // Export the chart to PDF directly into the memory stream
            chart.ToPdf(pdfStream);

            // Reset the stream position if you need to read from it later
            pdfStream.Position = 0;

            // Example: retrieve the PDF bytes for further in‑memory processing
            byte[] pdfBytes = pdfStream.ToArray();
            Console.WriteLine($"Chart PDF generated in memory. Size: {pdfBytes.Length} bytes");
        }
    }
}
