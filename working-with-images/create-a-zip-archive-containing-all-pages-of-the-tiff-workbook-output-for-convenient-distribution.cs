// Title: Generate one‑page TIFF images for each worksheet in an Aspose.Cells workbook and package them into a ZIP file using C#
// AI Prompts: Write C# code that iterates through all worksheets in a Workbook and uses Aspose.Cells SheetRender with ImageOrPrintOptions to save each sheet as a single‑page TIFF file. | Add logic to gather the generated TIFF files and create a ZIP archive with System.IO.Compression, preserving the original file names. | Include error handling that logs rendering failures per sheet, ensures temporary TIFF files are deleted after the ZIP is created, and returns the ZIP archive path.
// Common Searches: how to export each Excel worksheet to a separate TIFF file with Aspose.Cells in C# | C# create zip file containing multiple TIFF images generated from a workbook | Aspose.Cells one page per sheet TIFF export and compress into archive | remove temporary image files after zipping TIFFs in .NET application
// Tags: Aspose.Cells worksheet to TIFF conversion C# | create zip archive from TIFF files System.IO.Compression | single‑page TIFF export per sheet Aspose.Cells | temporary directory cleanup after image export .NET | error handling for sheet rendering Aspose.Cells

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook with three sheets, renders each sheet to a single‑page TIFF using Aspose.Cells SheetRender with specified resolution, stores the TIFFs in a temporary folder, compresses all TIFF files into a ZIP archive via System.IO.Compression, deletes the temporary folder, and outputs the path of the generated ZIP file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data to three worksheets
            Workbook workbook = new Workbook();
            for (int i = 0; i < 3; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                sheet.Name = $"Sheet{i + 1}";
                sheet.Cells["A1"].PutValue($"Data for {sheet.Name}");
                sheet.Cells["A2"].PutValue(100 + i);
            }

            // Folder to store temporary TIFF files (one per worksheet)
            string tempFolder = Path.Combine(Path.GetTempPath(), "TiffPages");
            Directory.CreateDirectory(tempFolder);

            // Render each worksheet to a single‑page TIFF file
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                try
                {
                    Worksheet sheet = workbook.Worksheets[i];

                    ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                    {
                        // ImageFormat property is optional; format is inferred from file extension
                        OnePagePerSheet = true,
                        HorizontalResolution = 200,
                        VerticalResolution = 200
                    };

                    SheetRender renderer = new SheetRender(sheet, imgOptions);
                    string tiffFile = Path.Combine(tempFolder, $"{sheet.Name}.tiff");
                    renderer.ToImage(0, tiffFile); // Save the first (and only) page
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error rendering sheet '{workbook.Worksheets[i].Name}': {ex.Message}");
                }
            }

            // Create a ZIP archive that contains all generated TIFF files
            string zipPath = Path.Combine(Environment.CurrentDirectory, "WorkbookPages.zip");
            try
            {
                if (File.Exists(zipPath))
                    File.Delete(zipPath);

                using (ZipArchive zip = ZipFile.Open(zipPath, ZipArchiveMode.Create))
                {
                    foreach (string tiffFile in Directory.GetFiles(tempFolder, "*.tiff"))
                    {
                        zip.CreateEntryFromFile(tiffFile, Path.GetFileName(tiffFile));
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating ZIP archive: {ex.Message}");
                return;
            }

            // Clean up temporary TIFF files
            try
            {
                if (Directory.Exists(tempFolder))
                    Directory.Delete(tempFolder, true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error cleaning temporary files: {ex.Message}");
            }

            Console.WriteLine($"ZIP archive created at: {zipPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
