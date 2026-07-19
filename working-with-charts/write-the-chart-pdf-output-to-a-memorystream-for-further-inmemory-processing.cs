// Title: Export Aspose.Cells Chart to PDF using a MemoryStream in C#
// Description: Creates a workbook, fills it with sample data, adds a column chart, and writes the chart directly to a MemoryStream with chart.ToPdf. The stream is reset and converted to a byte array for further in‑memory processing, eliminating the need for a temporary file.
// Keywords: Aspose.Cells chart PDF | MemoryStream export C# | chart.ToPdf Aspose | in‑memory PDF generation | CustomImplementationFactory | C# Excel chart to PDF | byte array PDF Aspose
// Common Searches: Aspose.Cells export chart to PDF memory stream | chart.ToPdf without saving file C# | generate PDF from Excel chart in memory | Aspose.Cells MemoryStream example | C# convert chart to PDF bytes
// Developer Intent: Produce a PDF representation of an Excel chart directly in memory for downstream processing, such as API responses or email attachments.
// Use Cases: Return chart PDF bytes from a web API without creating a file on disk | Attach an in‑memory chart PDF to an email using a byte array | Combine the chart PDF with other PDFs using a PDF merger library
// AI Prompts: Write C# code that builds an Aspose.Cells chart and exports it to a MemoryStream as PDF. | Explain why resetting the MemoryStream position after chart.ToPdf is required and show the correct code. | Demonstrate how to send the PDF bytes from the MemoryStream as an HTTP response in ASP.NET Core.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills it with sample data, adds a column chart, and writes the chart directly to a MemoryStream with chart.ToPdf. The stream is reset and converted to a byte array for further in‑memory processing, eliminating the need for a temporary file.
class ChartPdfToMemoryStreamDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["A3"].PutValue("Banana");
        worksheet.Cells["A4"].PutValue("Cherry");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["B3"].PutValue(45);
        worksheet.Cells["B4"].PutValue(25);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Create a MemoryStream using Aspose's factory method
        CustomImplementationFactory factory = new CustomImplementationFactory();
        using (MemoryStream pdfStream = factory.CreateMemoryStream())
        {
            // Export the chart to PDF directly into the memory stream
            chart.ToPdf(pdfStream);

            // Reset the stream position if further reading is required
            pdfStream.Position = 0;

            // Example: obtain the PDF bytes for further in‑memory processing
            byte[] pdfBytes = pdfStream.ToArray();
            Console.WriteLine($"Chart PDF generated in memory. Size: {pdfBytes.Length} bytes");
        }
    }
}
