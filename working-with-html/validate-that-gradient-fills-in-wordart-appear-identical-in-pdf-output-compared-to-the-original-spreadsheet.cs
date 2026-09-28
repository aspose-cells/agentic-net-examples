// Title: Check that WordArt gradient fills stay the same after converting an Excel sheet to PDF with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook, renders a chosen worksheet to a high‑resolution PNG, saves the workbook as PDF, reloads the PDF, renders the same page to PNG, and performs a pixel‑by‑pixel comparison to confirm WordArt gradient fills are unchanged. | Adapt the example to loop through every worksheet in the workbook, generate before‑and‑after PNGs for each, and produce a log that lists worksheets where gradient fill differences are detected after PDF conversion. | Enhance the comparison routine to create a diff image that highlights mismatched pixels between the original worksheet PNG and the PDF‑rendered PNG, using System.Drawing or a similar library.
// Common Searches: how to compare WordArt gradient colors in Excel and PDF using Aspose.Cells C# | Aspose.Cells render worksheet to PNG then to PDF and verify visual fidelity | C# pixel‑by‑pixel image comparison after Excel to PDF conversion | detect gradient fill differences in WordArt after saving Excel as PDF
// Tags: Aspose.Cells render worksheet to PNG | Aspose.Cells save workbook as PDF | compare Excel PNG with PDF PNG | WordArt gradient fill verification | pixel‑level image comparison C# | visual fidelity check after Excel to PDF conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace GradientFillComparison
{
    // The program loads an Excel workbook, renders the first worksheet to a PNG image, saves the workbook as PDF, reloads the PDF as a workbook, renders its first page back to PNG, and then compares the two PNG files byte‑by‑byte to determine whether WordArt gradient fills remain identical after conversion.
    class Program
    {
        static void Main(string[] args)
        {
            // Define file paths (replace placeholders with actual paths)
            string workbookPath = "{WorkbookPath}";
            string imagePathOriginal = "{ImagePathOriginal}";
            string pdfPath = "{PdfPath}";
            string imagePathPdf = "{ImagePathPdf}";

            try
            {
                // Verify the source workbook exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook not found: {workbookPath}");
                    return;
                }

                // Load the source Excel workbook
                var workbook = new Workbook(workbookPath);
                var worksheet = workbook.Worksheets[0];

                // Render the worksheet to an image (PNG) to capture the original appearance
                var imgOptions = new ImageOrPrintOptions
                {
                    // Default image format is PNG; no need to set ImageFormat explicitly
                    OnePagePerSheet = true,
                    Transparent = false,
                    HorizontalResolution = 300,
                    VerticalResolution = 300
                };
                var sheetRender = new SheetRender(worksheet, imgOptions);
                sheetRender.ToImage(0, imagePathOriginal);

                // Save the workbook as PDF
                workbook.Save(pdfPath, SaveFormat.Pdf);

                // Verify the generated PDF exists before loading
                if (!File.Exists(pdfPath))
                {
                    Console.WriteLine($"PDF not found after saving: {pdfPath}");
                    return;
                }

                // Render the first page of the generated PDF back to an image using Aspose.Cells
                var pdfWorkbook = new Workbook(pdfPath);
                var pdfWorksheet = pdfWorkbook.Worksheets[0];
                var pdfImgOptions = new ImageOrPrintOptions
                {
                    OnePagePerSheet = true,
                    Transparent = false,
                    HorizontalResolution = 300,
                    VerticalResolution = 300
                };
                var pdfSheetRender = new SheetRender(pdfWorksheet, pdfImgOptions);
                pdfSheetRender.ToImage(0, imagePathPdf);

                // Compare both PNG files byte‑by‑byte
                if (!File.Exists(imagePathOriginal) || !File.Exists(imagePathPdf))
                {
                    Console.WriteLine("One or both image files are missing; cannot compare.");
                    return;
                }

                byte[] originalBytes = File.ReadAllBytes(imagePathOriginal);
                byte[] pdfBytes = File.ReadAllBytes(imagePathPdf);

                bool identical = originalBytes.Length == pdfBytes.Length;
                if (identical)
                {
                    for (int i = 0; i < originalBytes.Length; i++)
                    {
                        if (originalBytes[i] != pdfBytes[i])
                        {
                            identical = false;
                            break;
                        }
                    }
                }

                Console.WriteLine(identical
                    ? "Gradient fills in WordArt are identical in PDF output."
                    : "Gradient fills differ between Excel and PDF.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
