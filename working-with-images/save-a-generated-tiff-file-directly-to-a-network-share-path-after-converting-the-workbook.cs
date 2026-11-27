// Title: Export an Aspose.Cells workbook to a 24‑bit LZW‑compressed TIFF and save it directly to a UNC network share using C#
// Common Searches: C# Aspose.Cells export worksheet to TIFF on a network share | How to write a TIFF file to a UNC path with Aspose.Cells .NET | Saving Aspose.Cells generated TIFF to a remote folder using SheetRender | Create directory and export workbook as TIFF to \\server\share in C# | Configure ImageOrPrintOptions for 300 DPI TIFF with LZW compression in Aspose.Cells
// Tags: Aspose.Cells export workbook to TIFF | SheetRender ToTiff network share | ImageOrPrintOptions 24-bit TIFF compression | UNC path file saving with Aspose.Cells | C# create directory before saving TIFF

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsTiffNetworkSave
{
    // The sample creates a workbook, applies styling, configures ImageOrPrintOptions for a 24‑bit LZW‑compressed TIFF at 300 DPI, and uses SheetRender.ToTiff to write the image directly to a UNC network share, creating the target directory if it does not exist.
    class Program
    {
        static void Main(string[] args)
        {
            // Define the UNC network share path where the TIFF will be saved.
            // Example: \\MyServer\SharedFolder\ConvertedWorkbook.tiff
            string networkTiffPath = @"\\MyServer\SharedFolder\ConvertedWorkbook.tiff";

            try
            {
                // Ensure the target directory exists. If it does not, attempt to create it.
                string targetDirectory = Path.GetDirectoryName(networkTiffPath);
                if (!string.IsNullOrEmpty(targetDirectory) && !Directory.Exists(targetDirectory))
                {
                    Console.WriteLine($"Target directory does not exist. Creating: {targetDirectory}");
                    Directory.CreateDirectory(targetDirectory);
                }

                // 1. Create a sample workbook (or load an existing one).
                Workbook workbook = new Workbook();

                // Populate the first worksheet with some sample data.
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "SampleData";

                sheet.Cells["A1"].PutValue("Aspose.Cells TIFF Export Demo");
                sheet.Cells["A2"].PutValue(DateTime.Now.ToString("F"));
                sheet.Cells["B4"].PutValue("Value 1");
                sheet.Cells["C4"].PutValue(12345);
                sheet.Cells["B5"].PutValue("Value 2");
                sheet.Cells["C5"].PutValue(67890);

                // Apply a simple style to demonstrate color depth handling.
                Style style = workbook.CreateStyle();
                style.Font.Color = System.Drawing.Color.DarkBlue;
                style.Font.Size = 14;
                style.Font.IsBold = true;
                style.ForegroundColor = System.Drawing.Color.LightYellow;
                style.Pattern = BackgroundType.Solid;
                sheet.Cells["A1"].SetStyle(style);

                // 2. Configure ImageOrPrintOptions for TIFF rendering.
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Tiff,                     // Output format
                    TiffCompression = TiffCompression.CompressionLZW, // LZW compression
                    TiffColorDepth = ColorDepth.Format24bpp,        // 24‑bit color depth
                    HorizontalResolution = 300,                     // DPI
                    VerticalResolution = 300,
                    OnePagePerSheet = true                          // Each sheet as a separate page
                };

                // 3. Render the worksheet to TIFF and save directly to the network path.
                SheetRender sheetRender = new SheetRender(sheet, imgOptions);
                sheetRender.ToTiff(networkTiffPath);

                Console.WriteLine($"TIFF file successfully saved to network share: {networkTiffPath}");
            }
            catch (Exception ex)
            {
                // Report any exception (e.g., access denied, path not found).
                Console.Error.WriteLine("An error occurred while saving the TIFF file:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
