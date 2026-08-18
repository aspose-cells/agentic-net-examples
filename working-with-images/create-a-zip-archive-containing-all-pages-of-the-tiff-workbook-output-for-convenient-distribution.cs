// Title: C# – Render Workbook Pages to Separate TIFF Files and Bundle Them into a ZIP Archive with Aspose.Cells
// Description: This example creates a workbook, fills it with sample data, configures ImageOrPrintOptions for one‑page‑per‑sheet TIFF output, renders each page using WorkbookRender, saves the TIFF files to a folder, and then compresses all images into a single ZIP file using System.IO.Compression. The original workbook is also saved for reference.
// Keywords: Aspose.Cells TIFF export C# | WorkbookRender one page per sheet | C# zip multiple TIFF files | export Excel pages as TIFF | .NET compress images to ZIP | Aspose.Cells image rendering
// Common Searches: how to export each Excel sheet page as TIFF with Aspose.Cells | C# render workbook pages to TIFF and zip them | Aspose.Cells OnePagePerSheet TIFF example | compress multiple TIFF images into a ZIP file in .NET | save Excel workbook as individual TIFF files
// Developer Intent: Generate a TIFF file for every workbook page and combine all TIFFs into a single ZIP archive.
// Use Cases: Distribute multi‑page Excel reports as a downloadable ZIP of per‑page TIFF images for universal viewing. | Archive printed versions of workbook pages for compliance, storing each page as a high‑resolution TIFF inside a compressed package. | Provide offline analysis of workbook content by bundling scanned pages as separate TIFF files within a ZIP file.
// AI Prompts: Write C# code that uses Aspose.Cells to render each workbook page to an individual TIFF file and then creates a ZIP archive of those files. | Explain how to automatically delete the temporary TIFF files after the ZIP archive has been created. | Show how to set DPI, compression type, and other TIFF options before adding the images to the ZIP archive with Aspose.Cells.

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsZipTiffExample
{
    // This example creates a workbook, fills it with sample data, configures ImageOrPrintOptions for one‑page‑per‑sheet TIFF output, renders each page using WorkbookRender, saves the TIFF files to a folder, and then compresses all images into a single ZIP file using System.IO.Compression. The original workbook is also saved for reference.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Aspose.Cells TIFF Pages to ZIP Demo");
            for (int i = 2; i <= 100; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Row {i - 1}");
            }

            // Configure image rendering options for TIFF output
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Tiff,          // Output format TIFF
                OnePagePerSheet = true               // Each page will be a separate image
            };

            // Initialize the workbook renderer
            WorkbookRender renderer = new WorkbookRender(workbook, imgOptions);

            // Prepare output directory
            string outputDir = "TiffPages";
            Directory.CreateDirectory(outputDir);

            // Render each page to an individual TIFF file
            for (int pageIndex = 0; pageIndex < renderer.PageCount; pageIndex++)
            {
                string pageFilePath = Path.Combine(outputDir, $"Page_{pageIndex + 1}.tiff");
                renderer.ToImage(pageIndex, pageFilePath);
            }

            // Create a ZIP archive containing all rendered TIFF pages
            string zipPath = "WorkbookPages.zip";
            using (FileStream zipStream = new FileStream(zipPath, FileMode.Create))
            using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Create))
            {
                foreach (string filePath in Directory.GetFiles(outputDir, "*.tiff"))
                {
                    // Add each TIFF file to the ZIP archive
                    archive.CreateEntryFromFile(filePath, Path.GetFileName(filePath));
                }
            }

            // Optional: clean up individual TIFF files after zipping
            // foreach (string file in Directory.GetFiles(outputDir, "*.tiff"))
            // {
            //     File.Delete(file);
            // }

            // Save the original workbook for reference
            workbook.Save("OriginalWorkbook.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("All pages rendered to TIFF, zipped into 'WorkbookPages.zip', and workbook saved.");
        }
    }
}
