// Title: Export an Excel timeline to PDF directly into a MemoryStream with Aspose.Cells for .NET
// Common Searches: how to save an Excel timeline as a PDF stream using Aspose.Cells C# | Aspose.Cells export timeline to PDF without creating a file | C# generate PDF from workbook containing a timeline into MemoryStream | in‑memory PDF conversion of Excel timeline with Aspose.Cells | save workbook timeline to PDF in a byte array using Aspose.Cells
// Tags: Aspose.Cells timeline PDF export MemoryStream | PdfSaveOptions for in‑memory PDF conversion | C# workbook to PDF stream without file | memory‑based PDF generation from Excel data | avoid disk I/O when exporting timeline to PDF

using System;
using System.IO;
using Aspose.Cells;

// The TimelinePdfExporter class creates a Workbook (assumed to already contain a timeline), configures PdfSaveOptions, and saves the workbook directly to a MemoryStream in PDF format. The stream position is reset before returning, enabling further processing without any disk I/O. The example also shows how to write the resulting stream to a file for verification, while handling errors via an ApplicationException.
public class TimelinePdfExporter
{
    public MemoryStream ExportTimelineToPdf()
    {
        try
        {
            // Create a new workbook (or load an existing one that contains a timeline)
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Add sample data, pivot table, and timeline here.
            // This section is omitted for brevity; assume the workbook
            // already contains the timeline you want to export.
            // -------------------------------------------------

            // Prepare a memory stream to hold the PDF output
            MemoryStream pdfStream = new MemoryStream();

            // Configure PDF save options if needed (e.g., page size, orientation)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook (including the timeline) directly to the memory stream in PDF format
            workbook.Save(pdfStream, pdfOptions);

            // Reset the stream position to the beginning for further processing
            pdfStream.Position = 0;

            // Return the memory stream containing the PDF data
            return pdfStream;
        }
        catch (Exception ex)
        {
            // Log or rethrow as needed; here we wrap in an ApplicationException for clarity
            throw new ApplicationException("Failed to export timeline to PDF.", ex);
        }
    }

    // Entry point for the console application
    public static void Main(string[] args)
    {
        try
        {
            TimelinePdfExporter exporter = new TimelinePdfExporter();
            using MemoryStream pdfStream = exporter.ExportTimelineToPdf();

            // Save the PDF to a file for verification
            string outputPath = "Timeline.pdf";
            File.WriteAllBytes(outputPath, pdfStream.ToArray());

            Console.WriteLine($"PDF successfully saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
