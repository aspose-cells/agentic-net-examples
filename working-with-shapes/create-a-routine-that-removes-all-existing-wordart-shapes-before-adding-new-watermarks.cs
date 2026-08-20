// Title: C# – Remove WordArt Shapes and Add PDF Watermark with Aspose.Cells
// Description: Loads an Excel workbook, deletes every WordArt shape from each worksheet, creates a red bold‑italic text watermark using RenderingFont and RenderingWatermark, and saves the workbook as a PDF with the watermark via PdfSaveOptions.
// Keywords: Aspose.Cells remove WordArt | C# delete WordArt shapes | Aspose.Cells PDF watermark | RenderingWatermark example | Excel to PDF with watermark | Aspose.Cells shape collection | remove WordArt before export | Aspose.Cells C# tutorial
// Common Searches: how to delete WordArt in Excel using Aspose.Cells | add text watermark when saving workbook to PDF Aspose.Cells | remove all WordArt shapes before PDF conversion C# | Aspose.Cells shape collection remove WordArt example | create diagonal CONFIDENTIAL watermark with Aspose.Cells
// Developer Intent: Strip all WordArt objects from a workbook and then apply a custom text watermark during PDF export.
// Use Cases: Prepare confidential reports by removing decorative WordArt and adding a semi‑transparent diagonal watermark. | Standardize batch conversion of Excel files to PDF where WordArt must be excluded for branding consistency. | Automate document sanitization for legal or compliance workflows, ensuring only the watermark remains visible in the final PDF.
// AI Prompts: Generate C# code that iterates through every worksheet in an Aspose.Cells workbook, removes shapes where IsWordArt is true, and saves the file as a PDF with a red, bold, italic watermark rotated 30°. | Show how to configure RenderingFont and RenderingWatermark for a 48‑pt, 25% opacity watermark using Aspose.Cells. | Explain how to modify the routine to preserve non‑WordArt shapes while still applying the PDF watermark.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace WatermarkApp
{
    // Loads an Excel workbook, deletes every WordArt shape from each worksheet, creates a red bold‑italic text watermark using RenderingFont and RenderingWatermark, and saves the workbook as a PDF with the watermark via PdfSaveOptions.
    public class WatermarkHelper
    {
        /// <param name="inputFile">Path to the source Excel file.</param>
        /// <param name="outputFile">Path where the PDF with watermark will be saved.</param>
        /// <param name="watermarkText">Text to be used for the watermark.</param>
        public static void RemoveWordArtAndAddWatermark(string inputFile, string outputFile, string watermarkText)
        {
            try
            {
                // Verify input file exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputFile);

                // Remove WordArt shapes from each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    ShapeCollection shapes = sheet.Shapes;
                    for (int i = shapes.Count - 1; i >= 0; i--)
                    {
                        Shape shape = shapes[i];
                        if (shape.IsWordArt)
                        {
                            shapes.RemoveAt(i);
                        }
                    }
                }

                // Create a rendering font for the watermark
                RenderingFont font = new RenderingFont("Arial", 48)
                {
                    Bold = true,
                    Italic = true,
                    Color = Color.Red
                };

                // Create the text watermark
                RenderingWatermark watermark = new RenderingWatermark(watermarkText, font)
                {
                    Rotation = 30,
                    Opacity = 0.25f,
                    IsBackground = true,
                    HAlignment = TextAlignmentType.Center,
                    VAlignment = TextAlignmentType.Center,
                    ScaleToPagePercent = 80
                };

                // Set the watermark in PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Watermark = watermark
                };

                // Save the workbook as PDF with the watermark
                workbook.Save(outputFile, pdfOptions);
                Console.WriteLine($"PDF saved with watermark to: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";
            string watermark = "CONFIDENTIAL";

            WatermarkHelper.RemoveWordArtAndAddWatermark(inputPath, outputPath, watermark);
        }
    }
}
