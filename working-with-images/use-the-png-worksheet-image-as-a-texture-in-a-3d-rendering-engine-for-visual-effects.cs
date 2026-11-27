// Title: Create a high‑resolution PNG of an Excel worksheet with Aspose.Cells in C# for use as a 3D texture
// Common Searches: C# Aspose.Cells render first worksheet to PNG with 300 DPI for Unity texture | How to convert Excel sheet to PNG image for 3D engine using Aspose.Cells | Generate high‑resolution spreadsheet image for texture mapping in C# | Aspose.Cells SheetRender export worksheet as PNG for Unreal Engine | Save Excel worksheet as PNG file programmatically in .NET Core
// Tags: Aspose.Cells SheetRender PNG export | C# Excel worksheet to high‑resolution image | render worksheet as texture bitmap | 300 DPI PNG generation from spreadsheet | Excel to PNG conversion for 3D rendering

using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTexture3D
{
    // The sample program ensures an input.xlsx workbook exists (creating a simple one if needed), then uses Aspose.Cells SheetRender with 300 dpi PNG settings to render the first worksheet to a high‑resolution sheet.png file, which can be loaded as a texture in 3D rendering engines.
    class Program
    {
        // Path to the source Excel file and the intermediate PNG image
        private const string ExcelFilePath = "input.xlsx";
        private const string PngFilePath = "sheet.png";

        static void Main(string[] args)
        {
            try
            {
                // ------------------------------------------------------------
                // Step 1 – Ensure we have an Excel file to work with.
                // If the file does not exist we create a simple workbook.
                // ------------------------------------------------------------
                if (!File.Exists(ExcelFilePath))
                {
                    Console.WriteLine($"Excel file \"{ExcelFilePath}\" not found. Creating a sample workbook.");
                    CreateSampleWorkbook(ExcelFilePath);
                }

                // ------------------------------------------------------------
                // Step 2 – Render the first worksheet to a PNG image.
                // ------------------------------------------------------------
                RenderWorksheetToPng(ExcelFilePath, PngFilePath);
                Console.WriteLine($"Worksheet rendered to PNG: {PngFilePath}");

                // ------------------------------------------------------------
                // Step 3 – Open the PNG with the default image viewer.
                // ------------------------------------------------------------
                if (File.Exists(PngFilePath))
                {
                    try
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = PngFilePath,
                            UseShellExecute = true
                        });
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Failed to open the PNG file:");
                        Console.WriteLine(ex.Message);
                    }
                }
                else
                {
                    Console.WriteLine($"PNG file \"{PngFilePath}\" was not created.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.ToString());
            }
        }

        // Creates a very small workbook with some data for demonstration.
        private static void CreateSampleWorkbook(string path)
        {
            var wb = new Workbook();
            var ws = wb.Worksheets[0];
            ws.Name = "Demo";

            ws.Cells["A1"].PutValue("Aspose.Cells");
            ws.Cells["A2"].PutValue("3D Texture Demo");
            ws.Cells["B1"].PutValue(DateTime.Now.ToString("yyyy-MM-dd"));
            ws.Cells["B2"].PutValue(12345);

            // Apply a simple style so the image looks nicer
            var style = wb.CreateStyle();
            style.Font.Name = "Arial";
            style.Font.Size = 14;
            style.Font.IsBold = true;
            ws.Cells["A1"].SetStyle(style);
            ws.Cells["A2"].SetStyle(style);

            wb.Save(path, SaveFormat.Xlsx);
        }

        // Uses Aspose.Cells SheetRender to create a PNG file from the first worksheet.
        private static void RenderWorksheetToPng(string excelPath, string pngPath)
        {
            // Load workbook
            var workbook = new Workbook(excelPath);
            var worksheet = workbook.Worksheets[0];

            // Configure rendering options
            var options = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png,
                OnePagePerSheet = true, // Render the whole sheet on a single page
                HorizontalResolution = 300,
                VerticalResolution = 300
            };

            // Render
            var sheetRender = new SheetRender(worksheet, options);

            // Ensure the output directory exists
            var dir = Path.GetDirectoryName(pngPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            // Render the first (and only) page to file
            sheetRender.ToImage(0, pngPath);
        }
    }
}
