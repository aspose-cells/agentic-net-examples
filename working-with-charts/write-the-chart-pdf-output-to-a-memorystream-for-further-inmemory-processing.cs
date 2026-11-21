// Title: Export an Aspose.Cells chart to PDF using a MemoryStream in C#
// Common Searches: C# how to save an Aspose.Cells workbook chart as a PDF in a MemoryStream | Aspose.Cells export chart to PDF stream without writing to disk | Using PdfSaveOptions with Aspose.Cells to generate an in‑memory PDF of a chart | Write chart PDF to MemoryStream for further processing in .NET | Aspose.Cells chart PDF generation to a byte array in C#
// Tags: Aspose.Cells chart PDF export via PdfSaveOptions | in‑memory PDF generation from workbook chart | C# MemoryStream handling for Aspose.Cells PDF output | column chart PDF rendering Aspose.Cells | resetting MemoryStream position after saving PDF

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The example creates a workbook, adds numeric data and a column chart, then uses PdfSaveOptions to save the workbook—including the chart—as a PDF directly into a MemoryStream. After resetting the stream position, the PDF bytes can be processed further or written to a file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and add some data
                Workbook workbook = new Workbook(); // create workbook
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["A3"].PutValue(30);

                // Add a column chart that uses the data
                // Worksheets.Charts.Add returns the index of the newly added chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("A1:A3", true);

                // Prepare a memory stream to hold the PDF output
                using (MemoryStream pdfStream = new MemoryStream())
                {
                    // Save the workbook (which contains the chart) as PDF into the memory stream
                    PdfSaveOptions pdfOptions = new PdfSaveOptions
                    {
                        OnePagePerSheet = true // optional: one page per sheet
                    };
                    workbook.Save(pdfStream, pdfOptions);

                    // Reset the stream position for further processing
                    pdfStream.Position = 0;

                    // Example: write PDF to a file (optional)
                    string outputPath = "output.pdf";
                    File.WriteAllBytes(outputPath, pdfStream.ToArray());
                    Console.WriteLine($"PDF saved to {Path.GetFullPath(outputPath)}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
