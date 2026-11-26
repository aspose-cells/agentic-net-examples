// Title: Create a single‑page TIFF from an Excel worksheet and overlay a PNG image using Aspose.Cells for .NET
// Common Searches: how to export an Excel worksheet to a single page TIFF with Aspose.Cells C# | add a PNG picture to an Excel sheet before converting to TIFF using Aspose.Cells | merge external raster image with workbook TIFF output in .NET | C# code to overlay a PNG on an Excel‑generated TIFF with Aspose.Cells | render worksheet to TIFF and combine with overlay image in Aspose.Cells
// Tags: Aspose.Cells worksheet to TIFF | C# add picture shape | overlay PNG on TIFF | single-page TIFF Aspose.Cells | merge external image with workbook TIFF

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTiffMerge
{
    // The sample loads or creates an Excel workbook, renders its first worksheet to a single‑page TIFF using ImageOrPrintOptions, optionally inserts a PNG picture as a shape on the sheet, and renders again to produce a merged TIFF; if the overlay image is missing, the original TIFF is copied as the final result.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output files
            string workbookPath = "SampleWorkbook.xlsx";          // Input Excel file (created if missing)
            string tiffPath = "WorkbookOutput.tiff";             // TIFF generated from the workbook
            string overlayImagePath = "OverlayImage.png";        // External raster image to merge
            string mergedOutputPath = "MergedResult.tiff";       // Final merged image

            try
            {
                // -------------------------------------------------
                // 1. Prepare a workbook (create if the file does not exist)
                // -------------------------------------------------
                Workbook workbook;
                if (File.Exists(workbookPath))
                {
                    workbook = new Workbook(workbookPath);
                }
                else
                {
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];
                    ws.Name = "Demo";
                    ws.Cells["A1"].PutValue("Aspose.Cells TIFF Merge Demo");
                    ws.Cells["A2"].PutValue(DateTime.Now);
                    ws.Cells["B2"].PutValue(12345);
                    workbook.Save(workbookPath, SaveFormat.Xlsx);
                }

                // -------------------------------------------------
                // 2. Render the workbook to a single-page TIFF image
                // -------------------------------------------------
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Tiff,
                    TiffCompression = TiffCompression.CompressionLZW,
                    HorizontalResolution = 300,
                    VerticalResolution = 300,
                    OnePagePerSheet = true
                };

                Worksheet firstSheet = workbook.Worksheets[0];
                SheetRender sheetRender = new SheetRender(firstSheet, renderOptions);
                sheetRender.ToTiff(tiffPath);
                Console.WriteLine($"Workbook rendered to TIFF: {tiffPath}");

                // -------------------------------------------------
                // 3. Merge overlay image if it exists
                // -------------------------------------------------
                if (!File.Exists(overlayImagePath))
                {
                    Console.WriteLine($"Overlay image not found at '{overlayImagePath}'. Skipping merge step.");
                    // No overlay – copy the original TIFF as the final result
                    File.Copy(tiffPath, mergedOutputPath, true);
                }
                else
                {
                    // Add the overlay image as a picture shape on the worksheet
                    // Position it at the top‑left cell (A1). Adjust as needed.
                    firstSheet.Pictures.Add(0, 0, overlayImagePath);

                    // Render again to obtain the merged image
                    SheetRender mergedRender = new SheetRender(firstSheet, renderOptions);
                    mergedRender.ToTiff(mergedOutputPath);
                    Console.WriteLine($"Merged image saved to: {mergedOutputPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
