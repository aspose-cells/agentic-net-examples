// Title: Create a Windows .ico icon from an Excel worksheet by rendering to SVG and PNG with Aspose.Cells for .NET (C#)
// Common Searches: how to create an .ico file from an Excel sheet using Aspose.Cells C# | render Excel worksheet to PNG and convert to Windows icon programmatically | Aspose.Cells generate high‑resolution icon from worksheet SVG example | C# code to build ICO file from PNG bytes with Aspose.Cells | temporary file handling when converting Excel to icon with Aspose.Cells
// Tags: Aspose.Cells SVG rendering of worksheet | Aspose.Cells high‑resolution PNG export | ICO file creation from PNG bytes C# | manual icon header and directory entry construction | temporary file management for Aspose.Cells rendering

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace SvgToIconConverter
{
    // The sample loads an Excel workbook, uses Aspose.Cells to render the first worksheet to an SVG file, then exports the same sheet as a 256 × 256 PNG at 300 dpi. It manually builds a Windows .ico file by writing the icon header, directory entry and embedding the PNG data, saves the .ico file, and cleans up the temporary SVG and PNG files.
    class Program
    {
        /// <param name="args">
        /// args[0] - path to the source Excel workbook (e.g., "sample.xlsx").
        /// args[1] - path where the generated icon should be saved (e.g., "app.ico").
        /// </param>
        static void Main(string[] args)
        {
            // Validate arguments
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: SvgToIconConverter <inputWorkbookPath> <outputIconPath>");
                return;
            }

            string workbookPath = args[0];
            string iconPath = args[1];

            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                return;
            }

            string svgTempPath = null;
            string pngTempPath = null;

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);
                Worksheet sheet = workbook.Worksheets[0]; // Use the first worksheet

                // -----------------------------------------------------------------
                // Step 1: Render the worksheet to SVG (intermediate file)
                // -----------------------------------------------------------------
                svgTempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".svg");
                ImageOrPrintOptions svgOptions = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Svg,
                    OnePagePerSheet = true
                };

                SheetRender svgRenderer = new SheetRender(sheet, svgOptions);
                using (MemoryStream svgStream = new MemoryStream())
                {
                    svgRenderer.ToImage(0, svgStream);
                    File.WriteAllBytes(svgTempPath, svgStream.ToArray());
                }
                Console.WriteLine($"SVG graphic saved to temporary file: {svgTempPath}");

                // -----------------------------------------------------------------
                // Step 2: Render the same worksheet to a high‑resolution PNG
                // -----------------------------------------------------------------
                const int iconSize = 256; // Desired icon size (pixels)
                pngTempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");

                ImageOrPrintOptions pngOptions = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Png,
                    OnePagePerSheet = true,
                    HorizontalResolution = 300,
                    VerticalResolution = 300
                };
                // Set the desired pixel dimensions
                pngOptions.SetDesiredSize(iconSize, iconSize, true);

                SheetRender pngRenderer = new SheetRender(sheet, pngOptions);
                using (MemoryStream pngStream = new MemoryStream())
                {
                    pngRenderer.ToImage(0, pngStream);
                    File.WriteAllBytes(pngTempPath, pngStream.ToArray());
                }
                Console.WriteLine($"PNG image saved to temporary file: {pngTempPath}");

                // -----------------------------------------------------------------
                // Step 3: Convert the PNG bitmap to an ICO file (manual ICO construction)
                // -----------------------------------------------------------------
                byte[] pngBytes = File.ReadAllBytes(pngTempPath);
                using (FileStream iconStream = new FileStream(iconPath, FileMode.Create, FileAccess.Write))
                {
                    // ICON header (6 bytes)
                    iconStream.WriteByte(0); // Reserved
                    iconStream.WriteByte(0);
                    iconStream.WriteByte(1); // Image type = 1 (icon)
                    iconStream.WriteByte(0);
                    iconStream.WriteByte(1); // Number of images
                    iconStream.WriteByte(0);

                    // Directory entry (16 bytes)
                    byte sizeByte = (byte)(iconSize == 256 ? 0 : iconSize); // 0 means 256 pixels
                    iconStream.WriteByte(sizeByte); // Width
                    iconStream.WriteByte(sizeByte); // Height
                    iconStream.WriteByte(0); // Color palette count (0 if no palette)
                    iconStream.WriteByte(0); // Reserved
                    // Color planes (2 bytes, little endian)
                    iconStream.WriteByte(1);
                    iconStream.WriteByte(0);
                    // Bits per pixel (2 bytes, little endian)
                    iconStream.WriteByte(32);
                    iconStream.WriteByte(0);
                    // Image data size (4 bytes, little endian)
                    byte[] sizeBytes = BitConverter.GetBytes(pngBytes.Length);
                    iconStream.Write(sizeBytes, 0, 4);
                    // Offset of image data from beginning of file (4 bytes, little endian)
                    int imageOffset = 6 + 16; // header + directory entry
                    byte[] offsetBytes = BitConverter.GetBytes(imageOffset);
                    iconStream.Write(offsetBytes, 0, 4);

                    // Image data (PNG bytes)
                    iconStream.Write(pngBytes, 0, pngBytes.Length);
                }

                Console.WriteLine($"Icon successfully created at: {iconPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during conversion:");
                Console.WriteLine(ex.Message);
            }
            finally
            {
                // Clean up temporary files if they exist
                try
                {
                    if (!string.IsNullOrEmpty(svgTempPath) && File.Exists(svgTempPath))
                        File.Delete(svgTempPath);
                    if (!string.IsNullOrEmpty(pngTempPath) && File.Exists(pngTempPath))
                        File.Delete(pngTempPath);
                }
                catch (Exception cleanupEx)
                {
                    Console.WriteLine("Warning: Failed to delete temporary files:");
                    Console.WriteLine(cleanupEx.Message);
                }
            }
        }
    }
}
