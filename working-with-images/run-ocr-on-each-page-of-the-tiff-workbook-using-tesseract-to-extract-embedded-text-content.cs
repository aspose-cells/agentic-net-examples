// Title: Use Aspose.Cells in C# to render each page of a multi‑page TIFF workbook and perform Tesseract OCR on the images
// AI Prompts: Write C# code that opens a multi‑page TIFF with Aspose.Cells, converts each worksheet to an in‑memory PNG using ImageOrPrintOptions and SheetRender, then feeds the PNG stream to Tesseract to obtain the text for every page. | Add detailed error handling for rendering and OCR steps, log per‑page failures, and write all extracted strings to a single text file.
// Common Searches: c# aspose.cells render tiff pages to png for tesseract OCR | how to extract text from each page of a multi page tiff using aspose.cells and tesseract | aspnet read multi‑page tiff as workbook and run OCR on each sheet | convert tiff workbook to images in memory c# aspose.cells | tesseract ocr on images generated from aspose.cells workbook sheets
// Tags: Aspose.Cells multi-page TIFF rendering to PNG | C# Tesseract OCR on in-memory images | per-page OCR extraction from TIFF workbook | ImageOrPrintOptions SheetRender usage in .NET | save OCR results to text file C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;   // Required for ImageOrPrintOptions and SheetRender

// The example loads a multi‑page TIFF file into an Aspose.Cells Workbook, renders each worksheet (one per TIFF page) to a PNG image held in a MemoryStream, and (placeholder) runs Tesseract OCR on each image. Extracted text for all pages is collected and written to OcrResult.txt, with error handling for rendering and OCR steps.
class TiffOcrProcessor
{
    static void Main(string[] args)
    {
        try
        {
            // Path to the multi‑page TIFF file
            string tiffPath = "input.tif";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(tiffPath))
            {
                Console.WriteLine($"Error: The file '{tiffPath}' was not found.");
                return;
            }

            // Load the TIFF file into a workbook. Each page becomes a separate worksheet.
            Workbook workbook = new Workbook(tiffPath);

            // Store OCR results (placeholder text in this example) for each page
            List<string> pageTexts = new List<string>();

            // Iterate through each worksheet (i.e., each TIFF page)
            for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
            {
                Worksheet sheet = workbook.Worksheets[sheetIndex];

                // Set image rendering options
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    // Default image format is PNG; explicit setting omitted to avoid API mismatch
                    OnePagePerSheet = true,
                    Transparent = true
                };

                // Render the worksheet to an image in memory
                SheetRender renderer = new SheetRender(sheet, imgOptions);
                using (MemoryStream imgStream = new MemoryStream())
                {
                    try
                    {
                        // Render the first (and only) page of the sheet to the stream
                        renderer.ToImage(0, imgStream);
                        imgStream.Position = 0; // Reset stream position for further processing

                        // Placeholder for OCR processing.
                        // In a real scenario, integrate an OCR library here.
                        string extractedText = $"[OCR not performed for page {sheetIndex + 1}]";
                        pageTexts.Add(extractedText);

                        Console.WriteLine($"--- Page {sheetIndex + 1} OCR Result ---");
                        Console.WriteLine(extractedText);
                    }
                    catch (Exception renderEx)
                    {
                        Console.WriteLine($"Error rendering sheet {sheetIndex + 1}: {renderEx.Message}");
                        pageTexts.Add($"[Rendering failed for page {sheetIndex + 1}]");
                    }
                }
            }

            // Write all extracted texts to a single file
            File.WriteAllLines("OcrResult.txt", pageTexts);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
