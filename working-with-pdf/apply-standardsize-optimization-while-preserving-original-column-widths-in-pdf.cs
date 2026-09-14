// Title: Export Excel to PDF while preserving column widths and disabling fit‑to‑page scaling with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx workbook using Aspose.Cells, disables automatic page scaling, and saves the workbook as a PDF. | Show how to configure PdfSaveOptions in Aspose.Cells to keep the original column layout when converting a worksheet to PDF. | Explain why the StandardSize optimization flag is unavailable in some Aspose.Cells releases and how to achieve a similar effect by turning off page scaling.
// Common Searches: keep original column layout when exporting Excel to PDF using Aspose.Cells | Aspose.Cells PdfSaveOptions disable fit to page example | export worksheet to PDF without scaling columns in C# | StandardSize PDF optimization missing in Aspose.Cells version | Excel to PDF conversion maintaining column widths Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions OnePagePerSheet false | preserve column widths Excel to PDF Aspose.Cells | PdfSaveOptions turn off page scaling C# | PDF optimization limitation Aspose.Cells | Excel workbook PDF conversion without column scaling

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel workbook (or creates a simple one if missing), sets PdfSaveOptions.OnePagePerSheet to false to prevent fit‑to‑page scaling, and saves the workbook as a PDF while preserving the original column widths using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            Workbook workbook;
            const string inputPath = "input.xlsx";

            // Load existing workbook if present; otherwise create a simple one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Cells["A1"].PutValue("Sample data");
            }

            // Configure PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Do not force fit‑to‑page; keep original column widths.
                OnePagePerSheet = false
                // Note: PdfOptimizationOptions is not available in this version of Aspose.Cells.
            };

            // Save the workbook as PDF.
            const string outputPath = "output.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
