// Title: C# batch conversion of XLSX files to 150 DPI PNG images with Aspose.Cells
// Description: Scans a given folder for *.xlsx workbooks, loads each with Aspose.Cells, sets CellsHelper.DPI to 150, and renders every worksheet page to a PNG file. Images are saved in a separate output directory using a naming pattern that includes the original workbook name, sheet index, and page number.
// Keywords: Aspose.Cells | C# XLSX to PNG | batch Excel image conversion | 150 DPI rendering | SheetRender PNG | convert multiple workbooks | Aspose.Cells example | GitHub Aspose.Cells PNG conversion
// Common Searches: C# convert all Excel files in a folder to PNG | Aspose.Cells batch render worksheets to PNG | set DPI for Excel to image conversion Aspose.Cells | export Excel worksheets as high‑resolution PNG | GitHub sample for XLSX to PNG batch conversion
// Developer Intent: Convert every XLSX file in a directory to PNG images at 150 DPI, producing one image per worksheet page and storing the results in a designated batch output folder.
// Use Cases: Archive Excel reports as printable PNG snapshots | Provide web‑ready previews of Excel dashboards | Feed Excel data into image‑processing or OCR pipelines | Create assets for documentation or e‑learning materials
// AI Prompts: Show how to change the DPI to 300 DPI while keeping the same folder structure. | Explain how to generate one PNG per worksheet without OnePagePerSheet, handling multi‑page sheets. | Add try‑catch logging and progress reporting to the batch converter. | Adapt the code for .NET Core and publish it as a GitHub Action.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace BatchXlsxToPng
{
    // Scans a given folder for *.xlsx workbooks, loads each with Aspose.Cells, sets CellsHelper.DPI to 150, and renders every worksheet page to a PNG file. Images are saved in a separate output directory using a naming pattern that includes the original workbook name, sheet index, and page number.
    class Program
    {
        static void Main()
        {
            // Folder containing source XLSX files
            string sourceFolder = @"C:\InputXlsx";
            // Folder where PNG images will be saved
            string outputFolder = @"C:\BatchPng";

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Set the desired DPI for rendering (150 DPI)
            CellsHelper.DPI = 150;

            // Get all XLSX files in the source folder
            string[] xlsxFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string xlsxPath in xlsxFiles)
            {
                // Load the workbook
                Workbook workbook = new Workbook(xlsxPath);

                // Iterate through each worksheet in the workbook
                for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
                {
                    Worksheet sheet = workbook.Worksheets[sheetIndex];

                    // Configure image rendering options for PNG
                    ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                    {
                        ImageType = ImageType.Png,
                        OnePagePerSheet = true   // Render the whole sheet on a single page
                    };

                    // Create a SheetRender instance for the current worksheet
                    SheetRender sheetRender = new SheetRender(sheet, imgOptions);

                    // Render each page of the sheet (usually one page because of OnePagePerSheet)
                    for (int page = 0; page < sheetRender.PageCount; page++)
                    {
                        // Build output file name: OriginalFileName_SheetIndex_Page.png
                        string fileName = Path.GetFileNameWithoutExtension(xlsxPath);
                        string outputPath = Path.Combine(
                            outputFolder,
                            $"{fileName}_Sheet{sheetIndex}_Page{page}.png");

                        // Save the rendered page directly to a PNG file
                        sheetRender.ToImage(page, outputPath);
                    }

                    // Release resources used by SheetRender
                    sheetRender.Dispose();
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
