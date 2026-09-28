// Title: Convert an Excel worksheet with embedded OLE objects to PDF while preserving placeholders using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx workbook containing OLE objects and saves it as a PDF with the OLE placeholders rendered as images using Aspose.Cells. | Show how to configure PdfSaveOptions for a PDF export when the ExportOleObjects property is not available in Aspose.Cells. | Add robust file‑existence verification and exception handling to a C# routine that converts an Excel file with OLE objects to PDF.
// Common Searches: how to export Excel sheet with OLE objects to PDF using Aspose.Cells C# | Aspose.Cells PDF conversion keep OLE placeholders as images | C# PdfSaveOptions default behavior for embedded OLE objects in Excel
// Tags: Aspose.Cells PDF export with OLE objects | C# preserve OLE placeholders during Excel to PDF conversion | PdfSaveOptions default handling of embedded OLE | file existence check Aspose.Cells workbook conversion | exception handling Aspose.Cells PDF save

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that InputWithOleObjects.xlsx exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions instance (the ExportOleObjects flag is unavailable, so OLE objects are rendered as images by default), and saves the workbook as OutputPreservingOleObjects.pdf. It includes basic error handling to report missing files or other exceptions.
class ExportOleObjectsToPdf
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWithOleObjects.xlsx";
            const string outputPath = "OutputPreservingOleObjects.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook that contains OLE objects.
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options. (ExportOleObjects property is not available in this version;
            // OLE objects will be rendered as images by default.)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook to PDF while preserving OLE placeholders.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
