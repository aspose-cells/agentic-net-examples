using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class ConvertXlsxToTiffWithProgress
{
    static void Main()
    {
        // Paths
        string sourcePath = "input.xlsx";
        string outputFolder = "output_tiff_pages";

        // Ensure output directory exists
        Directory.CreateDirectory(outputFolder);

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(sourcePath);

        // Configure image options for TIFF output
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.SaveFormat = SaveFormat.Tiff;      // Set output format to TIFF
        options.OnePagePerSheet = true;            // Render each page separately

        // Create a renderer for the workbook
        WorkbookRender renderer = new WorkbookRender(workbook, options);
        int totalPages = renderer.PageCount;

        Console.WriteLine($"Total pages to render: {totalPages}");

        // Render each page and provide progress feedback
        for (int i = 0; i < totalPages; i++)
        {
            Console.WriteLine($"Rendering page {i + 1} of {totalPages}...");

            // Render the current page to a memory stream
            using (MemoryStream ms = new MemoryStream())
            {
                renderer.ToImage(i, ms); // Render page i to the stream

                // Save the stream content as a TIFF file
                string outPath = Path.Combine(outputFolder, $"page_{i + 1}.tiff");
                File.WriteAllBytes(outPath, ms.ToArray());
            }

            Console.WriteLine($"Page {i + 1} saved.");
        }

        Console.WriteLine("Conversion completed successfully.");
    }
}