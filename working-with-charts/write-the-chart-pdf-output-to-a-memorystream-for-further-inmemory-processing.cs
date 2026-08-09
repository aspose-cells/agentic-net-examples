// Title: Export Aspose.Cells Chart to PDF via MemoryStream (C#) – In‑Memory Processing
// Description: Demonstrates how to create a workbook, add a column chart, and write the chart directly to a PDF stored in a MemoryStream using Aspose.Cells. The stream can be reset, read as a byte array, or passed to other services without creating a temporary file.
// Keywords: Aspose.Cells chart to PDF | C# MemoryStream PDF export | export Excel chart in memory | Aspose.Cells ToPdf MemoryStream | in‑memory PDF generation | chart PDF byte array | Aspose.Cells API example
// Common Searches: Aspose.Cells export chart to PDF memory stream C# | How to save chart as PDF in memory using Aspose | C# write Aspose chart to byte array | In‑memory PDF from Excel chart Aspose.Cells | Aspose.Cells ToPdf without file
// Developer Intent: Create a PDF representation of an Excel chart directly in a MemoryStream for downstream processing such as API responses, database storage, or PDF merging.
// Use Cases: Return the chart PDF bytes from a Web API without touching the file system. | Store the generated PDF in a database BLOB for later retrieval. | Combine several chart PDFs held in MemoryStreams into a single document with a PDF library.
// AI Prompts: Generate C# code that exports an Aspose.Cells chart to a MemoryStream and attaches the resulting PDF to an email. | Show how to merge multiple chart PDFs created in MemoryStreams into one PDF using Aspose.Pdf. | Explain how to set custom page size and orientation when exporting an Aspose.Cells chart to a MemoryStream.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartPdfToMemoryStream
{
    // Demonstrates how to create a workbook, add a column chart, and write the chart directly to a PDF stored in a MemoryStream using Aspose.Cells. The stream can be reset, read as a byte array, or passed to other services without creating a temporary file.
    public class ChartPdfMemoryStreamDemo
    {
        public static void Run()
        {
            try
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

                // Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Create a MemoryStream using the provided factory rule
                CustomImplementationFactory factory = new CustomImplementationFactory();
                using (MemoryStream pdfStream = factory.CreateMemoryStream())
                {
                    // Export the chart to PDF directly into the memory stream
                    chart.ToPdf(pdfStream);

                    // Reset the stream position if you need to read from it later
                    pdfStream.Position = 0;

                    // Example of further in‑memory processing: get the PDF bytes
                    byte[] pdfBytes = pdfStream.ToArray();
                    Console.WriteLine($"PDF generated in memory. Size: {pdfBytes.Length} bytes");

                    // (Optional) Write the PDF to a file for verification
                    // File.WriteAllBytes("ChartOutput.pdf", pdfBytes);
                }

                // The workbook itself can also be saved to a memory stream if needed
                // using (MemoryStream wbStream = workbook.SaveToStream()) { ... }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Simple factory to create MemoryStream instances (placeholder for actual implementation)
    public class CustomImplementationFactory
    {
        public MemoryStream CreateMemoryStream()
        {
            return new MemoryStream();
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ChartPdfMemoryStreamDemo.Run();
        }
    }
}
