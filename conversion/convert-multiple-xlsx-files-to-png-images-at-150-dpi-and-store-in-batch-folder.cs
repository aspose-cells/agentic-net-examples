// Title: How to batch convert multiple XLSX workbooks to 150 DPI PNG images using Aspose.Cells in C#
// AI Prompts: Generate a C# console program that scans a directory for .xlsx files, sets CellsHelper.DPI to 150, and saves each worksheet as a separate PNG file using Aspose.Cells. | Write .NET code to convert all Excel workbooks in a directory into high‑resolution PNG images, generating one image per sheet and applying a custom DPI. | Create a C# script that creates an output folder, iterates over XLSX files, and uses WorkbookRender with ImageOrPrintOptions to produce PNG files at 150 DPI.
// Common Searches: C# Aspose.Cells batch export Excel worksheets to PNG with specific DPI | How to set rendering DPI when converting XLSX to PNG using Aspose.Cells | Render each sheet of multiple Excel files to separate PNG files in .NET | Automate conversion of a folder of .xlsx files to high‑resolution PNG images
// Tags: Aspose.Cells batch XLSX to PNG conversion | CellsHelper DPI configuration for image rendering | WorkbookRender export worksheets as PNG | ImageOrPrintOptions PNG one page per sheet | C# automate Excel to high‑resolution image

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace BatchXlsxToPng
{
    // The C# console app sets CellsHelper.DPI to 150, creates an output directory, iterates over every .xlsx file in a source folder, loads each workbook with Aspose.Cells, and uses WorkbookRender together with ImageOrPrintOptions (PNG format, one page per sheet) to render each worksheet to a separate PNG file named with the original workbook and page index.
    class Program
    {
        static void Main()
        {
            // Set the desired DPI for rendering
            CellsHelper.DPI = 150;

            // Folder containing source XLSX files
            string sourceFolder = @"C:\InputXlsx";

            // Folder where PNG images will be saved
            string outputFolder = @"C:\BatchPng";
            Directory.CreateDirectory(outputFolder);

            // Image rendering options (PNG format, one page per sheet)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                OnePagePerSheet = true
            };

            // Process each XLSX file in the source folder
            foreach (string xlsxPath in Directory.GetFiles(sourceFolder, "*.xlsx"))
            {
                // Load the workbook
                Workbook workbook = new Workbook(xlsxPath);

                // Create a renderer for the whole workbook
                WorkbookRender renderer = new WorkbookRender(workbook, imgOptions);

                // Render each page (sheet) to a separate PNG file
                for (int pageIndex = 0; pageIndex < renderer.PageCount; pageIndex++)
                {
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsxPath);
                    string pngPath = Path.Combine(outputFolder,
                        $"{fileNameWithoutExt}_page{pageIndex}.png");

                    // Render the page to the specified PNG file
                    renderer.ToImage(pageIndex, pngPath);
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
