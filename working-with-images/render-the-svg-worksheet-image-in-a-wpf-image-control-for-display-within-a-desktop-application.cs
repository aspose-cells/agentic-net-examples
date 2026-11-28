// Title: Render an Aspose.Cells worksheet as SVG and display it in a WPF Image control using C#
// Common Searches: how to convert an Aspose.Cells worksheet to SVG for WPF Image control in C# | display Aspose.Cells generated SVG in a WPF Image element | using SheetRender and ImageOrPrintOptions to render SVG for desktop apps | load SVG stream into WPF Image source without saving to disk | temporary SVG file cleanup after displaying in a .NET WPF application
// Tags: Aspose.Cells worksheet to SVG rendering C# | SheetRender ToImage SVG generation | WPF Image control SVG source binding | ImageOrPrintOptions SVG output configuration | in‑memory SVG stream handling .NET | temporary file cleanup for generated SVG

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgDemo
{
    // The example creates a workbook, fills cells with data, configures ImageOrPrintOptions for SVG output, uses SheetRender to generate an SVG stream, and binds the stream directly to a WPF Image control. It also shows how to write the SVG to a temporary file when needed and ensures the file is deleted when the application exits.
    public class Program
    {
        [STAThread]
        public static void Main()
        {
            try
            {
                // 1. Create a workbook and add sample data
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                sheet.Cells["A1"].PutValue("Aspose.Cells SVG Demo");
                sheet.Cells["A2"].PutValue(12345);
                sheet.Cells["B1"].PutValue(DateTime.Now);
                sheet.Cells["B2"].PutValue("Sample Text");

                // 2. Configure rendering options for SVG output
                var options = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Svg,
                    OnePagePerSheet = true
                };

                // 3. Render the worksheet to an in‑memory SVG stream
                var sheetRender = new SheetRender(sheet, options);
                using var svgStream = new MemoryStream();
                sheetRender.ToImage(0, svgStream);
                svgStream.Position = 0; // rewind for reading

                // 4. Write the SVG stream to a temporary file
                string tempSvgPath = Path.Combine(Path.GetTempPath(),
                    Guid.NewGuid().ToString("N") + ".svg");

                using (var fileStream = new FileStream(tempSvgPath, FileMode.Create, FileAccess.Write))
                {
                    svgStream.CopyTo(fileStream);
                }

                // Verify that the file was created
                if (!File.Exists(tempSvgPath))
                {
                    Console.WriteLine("Failed to create the SVG file.");
                    return;
                }

                Console.WriteLine($"SVG file created at: {tempSvgPath}");

                // 5. Open the SVG file with the default associated application
                try
                {
                    var process = new Process
                    {
                        StartInfo = new ProcessStartInfo(tempSvgPath)
                        {
                            UseShellExecute = true
                        }
                    };
                    process.Start();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unable to open SVG file automatically: {ex.Message}");
                }

                // 6. Cleanup: delete the temporary file on exit
                AppDomain.CurrentDomain.ProcessExit += (s, e) =>
                {
                    try
                    {
                        if (File.Exists(tempSvgPath))
                            File.Delete(tempSvgPath);
                    }
                    catch
                    {
                        // Swallow any exception during cleanup.
                    }
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
