// Title: Compare LZW and CCITT4 TIFF compression ratios with Aspose.Cells in C#
// Common Searches: C# code to render an Excel workbook to TIFF using LZW compression with Aspose.Cells | How to measure file size difference between LZW and CCITT4 TIFF in Aspose.Cells | Aspose.Cells C# example for comparing TIFF compression options | Determine optimal TIFF compression (LZW vs CCITT4) for Excel export in .NET | Benchmark TIFF output size using Aspose.Cells WorkbookRender compression settings
// Tags: render workbook to TIFF with LZW compression Aspose.Cells | render workbook to TIFF with CCITT4 compression Aspose.Cells | compare TIFF compression methods Aspose.Cells C# | tiff compression ratio analysis Aspose.Cells | image size reduction using Aspose.Cells rendering options

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace TiffCompressionComparison
{
    // The sample creates a workbook with data and a chart, renders it twice to TIFF—once with LZW compression and once with CCITT4 compression—captures each file's size, and reports which compression yields the smaller file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a sample workbook with enough data to make the test meaningful
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate the worksheet with mixed text, numbers and a simple chart
                for (int row = 0; row < 100; row++)
                {
                    for (int col = 0; col < 10; col++)
                    {
                        sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Add a simple column chart (charts are also rendered into the TIFF)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 105, 0, 130, 5);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("A1:J100", true);
                chart.Title.Text = "Sample Chart";

                // -----------------------------------------------------------------
                // 2. Prepare common image options (TIFF output)
                // -----------------------------------------------------------------
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Tiff,
                    HorizontalResolution = 300,
                    VerticalResolution = 300,
                    OnePagePerSheet = true // Required for TIFF output
                };

                // -----------------------------------------------------------------
                // 3. Render with LZW compression
                // -----------------------------------------------------------------
                options.TiffCompression = TiffCompression.CompressionLZW;
                string lzwPath = Path.Combine(Path.GetTempPath(), "Workbook_LZW.tiff");
                WorkbookRender lzwRender = new WorkbookRender(workbook, options);
                lzwRender.ToImage(lzwPath);
                long lzwSize = new FileInfo(lzwPath).Length;
                Console.WriteLine($"LZW compressed TIFF saved to: {lzwPath}");
                Console.WriteLine($"File size (LZW): {lzwSize} bytes");

                // -----------------------------------------------------------------
                // 4. Render with CCITT4 compression
                // -----------------------------------------------------------------
                // Note: CCITT4 works only with 1‑bit (black‑white) images.
                // The renderer will automatically convert the page to 1‑bpp.
                options.TiffCompression = TiffCompression.CompressionCCITT4;
                string ccitt4Path = Path.Combine(Path.GetTempPath(), "Workbook_CCITT4.tiff");
                WorkbookRender ccitt4Render = new WorkbookRender(workbook, options);
                ccitt4Render.ToImage(ccitt4Path);
                long ccitt4Size = new FileInfo(ccitt4Path).Length;
                Console.WriteLine($"CCITT4 compressed TIFF saved to: {ccitt4Path}");
                Console.WriteLine($"File size (CCITT4): {ccitt4Size} bytes");

                // -----------------------------------------------------------------
                // 5. Compare the two compression results
                // -----------------------------------------------------------------
                if (lzwSize < ccitt4Size)
                {
                    Console.WriteLine("LZW compression yields a smaller file.");
                }
                else if (ccitt4Size < lzwSize)
                {
                    Console.WriteLine("CCITT4 compression yields a smaller file.");
                }
                else
                {
                    Console.WriteLine("Both compression methods produced files of identical size.");
                }

                // -----------------------------------------------------------------
                // 6. Optional cleanup (comment out if you want to inspect the files)
                // -----------------------------------------------------------------
                // File.Delete(lzwPath);
                // File.Delete(ccitt4Path);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
