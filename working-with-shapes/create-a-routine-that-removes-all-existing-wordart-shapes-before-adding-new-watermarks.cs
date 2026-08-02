// Title: C# – Remove all WordArt shapes before applying a PDF watermark with Aspose.Cells
// Description: Loads an Excel workbook, iterates each worksheet, deletes every shape whose IsWordArt flag is true, then creates a RenderingFont and RenderingWatermark, configures PdfSaveOptions, and saves the workbook as a PDF with a semi‑transparent, rotated watermark.
// Keywords: Aspose.Cells remove WordArt | delete WordArt shapes C# | Aspose.Cells PDF watermark | shape collection remove WordArt | Excel to PDF watermark Aspose | C# Aspose.Cells rendering watermark
// Common Searches: how to delete WordArt in Aspose.Cells before PDF export | Aspose.Cells remove WordArt shapes C# example | apply PDF watermark after clearing WordArt with Aspose.Cells | C# code to clear all WordArt from Excel workbook using Aspose | Aspose.Cells remove decorative shapes before saving as PDF
// Developer Intent: Delete every WordArt shape in all worksheets, then add a configurable PDF watermark using Aspose.Cells for .NET.
// Use Cases: Prepare confidential reports by stripping WordArt and overlaying a semi‑transparent watermark. | Clean up client workbooks before batch conversion to PDF to avoid visual clutter. | Automate corporate branding enforcement by removing decorative shapes and applying a standard watermark.
// AI Prompts: Generate C# code that iterates through all worksheets in an Aspose.Cells workbook, removes shapes where IsWordArt is true, and saves the file as a PDF with a custom watermark. | Show an alternative using ShapeCollection.RemoveAll with a lambda to delete WordArt shapes in Aspose.Cells. | Explain how to adjust watermark opacity, rotation, alignment, and scaling while ensuring WordArt shapes are removed first.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkDemo
{
    // Loads an Excel workbook, iterates each worksheet, deletes every shape whose IsWordArt flag is true, then creates a RenderingFont and RenderingWatermark, configures PdfSaveOptions, and saves the workbook as a PDF with a semi‑transparent, rotated watermark.
    public class WatermarkHelper
    {
        /// <param name="inputFile">Path to the source Excel file.</param>
        /// <param name="outputFile">Path where the watermarked PDF will be saved.</param>
        /// <param name="watermarkText">Text to be used for the watermark.</param>
        public static void ApplyWatermark(string inputFile, string outputFile, string watermarkText)
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
                throw new FileNotFoundException($"Input file not found: {inputFile}");

            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputFile);

                // Iterate through all worksheets and remove existing WordArt shapes
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    ShapeCollection shapes = sheet.Shapes;

                    // Collect indices of WordArt shapes (IsWordArt == true)
                    var wordArtIndices = new System.Collections.Generic.List<int>();
                    for (int i = 0; i < shapes.Count; i++)
                    {
                        Shape shape = shapes[i];
                        if (shape.IsWordArt)
                            wordArtIndices.Add(i);
                    }

                    // Remove WordArt shapes in reverse order to keep indices valid
                    for (int i = wordArtIndices.Count - 1; i >= 0; i--)
                        shapes.RemoveAt(wordArtIndices[i]);
                }

                // Create a rendering font for the watermark
                RenderingFont font = new RenderingFont("Arial", 48)
                {
                    Bold = true,
                    Color = Color.Red,
                    Italic = false
                };

                // Configure the text watermark
                RenderingWatermark watermark = new RenderingWatermark(watermarkText, font)
                {
                    Rotation = 45f,
                    Opacity = 0.3f,
                    IsBackground = true,
                    HAlignment = TextAlignmentType.Center,
                    VAlignment = TextAlignmentType.Center,
                    ScaleToPagePercent = 70
                };

                // Set the watermark in PDF save options
                PdfSaveOptions saveOptions = new PdfSaveOptions
                {
                    Watermark = watermark
                };

                // Save the workbook as PDF with the watermark applied
                workbook.Save(outputFile, saveOptions);
            }
            catch (Exception ex)
            {
                // Rethrow with additional context
                throw new ApplicationException($"Failed to apply watermark to '{inputFile}'.", ex);
            }
        }

        // Example usage
        public static void Main()
        {
            string sourceExcel = "InputWorkbook.xlsx";
            string resultPdf = "WatermarkedOutput.pdf";
            string text = "CONFIDENTIAL";

            try
            {
                ApplyWatermark(sourceExcel, resultPdf, text);
                Console.WriteLine("Watermark applied and PDF saved to: " + resultPdf);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
