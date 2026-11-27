// Title: C# example: Convert an Excel workbook with WordArt to PDF/A‑2b using Aspose.Cells while preserving gradient fills
// Common Searches: c# asp.net convert excel file that contains wordart to pdf/a-2b preserving gradients | aspose.cells pdfsaveoptions pdfa2b example with wordart shapes | how to detect wordart objects in a workbook before exporting to pdf with aspose.cells | export excel to pdf/a-2b in .net ensuring shape formatting is retained | save workbook as pdf/a-2b without flattening gradient fills using aspose.cells c#
// Tags: Aspose.Cells PDF/A-2b export with WordArt support | C# PdfSaveOptions compliance PdfA2b gradient preservation | WordArt shape detection Aspose.Cells API | Excel to PDF/A-2b conversion .NET | preserve gradient fills during Excel PDF export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;          // Required for Shape and ShapeCollection
using Aspose.Cells.Rendering;        // Required for PdfSaveOptions and PdfCompliance

namespace WordArtToPdfA2b
{
    // The sample loads an Excel workbook, enumerates worksheets to identify WordArt shapes, configures PdfSaveOptions for PDF/A‑2b compliance, and saves the file as a PDF/A‑2b document while keeping gradient fills and WordArt rendering intact.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "WordArtWorkbook.xlsx";
            const string outputPath = "WordArtOutput_Pdfa2b.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook that contains WordArt shapes
                Workbook workbook = new Workbook(inputPath);
                Console.WriteLine($"Loaded workbook: {inputPath}");

                // Enumerate worksheets and report WordArt shapes for verification
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    ShapeCollection shapes = sheet.Shapes;
                    for (int i = 0; i < shapes.Count; i++)
                    {
                        Shape shape = shapes[i];
                        if (shape.IsWordArt)
                        {
                            Console.WriteLine($"Found WordArt shape on sheet \"{sheet.Name}\" at index {i} with text: \"{shape.Text}\"");
                        }
                    }
                }

                // Configure PDF save options for PDF/A‑2b compliance
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Compliance = PdfCompliance.PdfA2b,
                    OnePagePerSheet = false
                };

                // Save the workbook as PDF/A‑2b
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF/A‑2b to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during conversion:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
