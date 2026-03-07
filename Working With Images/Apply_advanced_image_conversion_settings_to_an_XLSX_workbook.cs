using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class AdvancedImageConversion
{
    static void Main()
    {
        // Input Excel file (XLSX) and output image files
        string sourceFile = "input.xlsx";
        string outputSinglePage = "output_page0.tiff";
        string outputMultiPage = "output_multi.tiff";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(sourceFile);

        // -------------------------------------------------
        // Configure advanced image rendering options
        // -------------------------------------------------
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
        imgOptions.ImageType = ImageType.Tiff;                     // Output format
        imgOptions.TiffCompression = TiffCompression.CompressionLZW; // LZW compression
        imgOptions.HorizontalResolution = 300;                    // 300 DPI horizontal
        imgOptions.VerticalResolution = 300;                      // 300 DPI vertical
        imgOptions.OnePagePerSheet = true;                        // One page per worksheet
        imgOptions.AllColumnsInOnePagePerSheet = true;            // Fit all columns on one page

        // -------------------------------------------------
        // Render the first page (sheet) to a single TIFF file using WorkbookRender
        // -------------------------------------------------
        WorkbookRender renderer = new WorkbookRender(workbook, imgOptions);
        // Render page index 0 (first sheet) to the specified file
        renderer.ToImage(0, outputSinglePage);

        // -------------------------------------------------
        // Save the entire workbook as a multi‑page TIFF using ImageSaveOptions
        // -------------------------------------------------
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // Apply the same rendering options to the save options
        saveOptions.ImageOrPrintOptions.ImageType = ImageType.Tiff;
        saveOptions.ImageOrPrintOptions.TiffCompression = TiffCompression.CompressionLZW;
        saveOptions.ImageOrPrintOptions.HorizontalResolution = 300;
        saveOptions.ImageOrPrintOptions.VerticalResolution = 300;
        saveOptions.ImageOrPrintOptions.OnePagePerSheet = true;
        saveOptions.ImageOrPrintOptions.AllColumnsInOnePagePerSheet = true;

        // Save the workbook as a multi‑page TIFF image
        workbook.Save(outputMultiPage, saveOptions);
    }
}