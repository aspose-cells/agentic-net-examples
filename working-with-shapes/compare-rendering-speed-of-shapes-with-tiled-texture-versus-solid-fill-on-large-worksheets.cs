// Title: Measure Aspose.Cells rendering performance of tiled texture vs solid fill shapes on a large Excel worksheet (C#)
// Common Searches: C# Aspose.Cells benchmark rendering time for worksheets with textured shapes | compare rendering speed of texture fill and solid fill shapes using Aspose.Cells SheetRender | how to test performance of shape fills in large Excel files with Aspose.Cells .NET | measure image generation time for a worksheet containing 400 shapes in Aspose.Cells | Aspose.Cells rendering performance with many shapes on a 2000‑row worksheet
// Tags: Aspose.Cells texture fill shape rendering benchmark | Aspose.Cells solid fill shape performance test | SheetRender rendering time measurement | large Excel worksheet shape rendering .NET | C# shape fill type performance analysis

using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace ShapeRenderSpeedComparison
{
    // The sample creates two workbooks each with 2,000 rows and 30 columns, adds 400 rectangle shapes per workbook, and applies either a tiled texture fill or a solid fill to each shape. It then renders the first worksheet of each workbook to a PNG image using SheetRender, records the rendering duration with a Stopwatch, outputs the timings, and saves both the images and the workbooks for further inspection.
    class Program
    {
        // Configuration
        private const int TotalRows = 2000;
        private const int TotalColumns = 30;
        private const int ShapeCount = 400; // Number of shapes to add for each test
        private const int ShapeHeight = 30;
        private const int ShapeWidth = 80;

        static void Main(string[] args)
        {
            try
            {
                // License handling (if you have a license file, uncomment and set the correct path)
                // var license = new License();
                // license.SetLicense("Aspose.Cells.lic");

                Console.WriteLine("Starting rendering speed comparison...");

                // 1. Workbook with tiled texture fill shapes
                var tiledWorkbook = CreateWorkbook(withTiledTexture: true);
                var tiledTime = RenderWorksheetToImage(tiledWorkbook, "TiledTexture.png");
                Console.WriteLine($"Tiled texture fill rendering time: {tiledTime.TotalMilliseconds} ms");

                // 2. Workbook with solid fill shapes
                var solidWorkbook = CreateWorkbook(withTiledTexture: false);
                var solidTime = RenderWorksheetToImage(solidWorkbook, "SolidFill.png");
                Console.WriteLine($"Solid fill rendering time: {solidTime.TotalMilliseconds} ms");

                // Save workbooks for manual inspection (optional)
                tiledWorkbook.Save("TiledTextureWorkbook.xlsx", SaveFormat.Xlsx);
                solidWorkbook.Save("SolidFillWorkbook.xlsx", SaveFormat.Xlsx);

                Console.WriteLine("Comparison completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }

        /// <param name="withTiledTexture">If true, shapes use a tiled texture fill; otherwise solid fill.</param>
        /// <returns>The prepared Workbook instance.</returns>
        private static Workbook CreateWorkbook(bool withTiledTexture)
        {
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Populate the worksheet with dummy data to simulate a large sheet
            for (int row = 0; row < TotalRows; row++)
            {
                for (int col = 0; col < TotalColumns; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // Add shapes
            var shapes = sheet.Shapes;
            var rnd = new Random();

            for (int i = 0; i < ShapeCount; i++)
            {
                // Random position within the used range
                int topRow = rnd.Next(0, TotalRows - 1);
                int leftColumn = rnd.Next(0, TotalColumns - 1);
                int topOffset = rnd.Next(0, 20);
                int leftOffset = rnd.Next(0, 20);

                // Add rectangle shape
                Shape shape = shapes.AddRectangle(topRow, topOffset, leftColumn, leftOffset, ShapeHeight, ShapeWidth);
                shape.Placement = PlacementType.MoveAndSize; // Ensure shape moves with cells

                if (withTiledTexture)
                {
                    // Apply tiled texture fill
                    shape.Fill.FillType = FillType.Texture;
                    var textureFill = shape.Fill.TextureFill;
                    textureFill.Type = TextureType.BlueTissuePaper; // Built‑in texture
                    textureFill.IsTiling = true;

                    // Optional: adjust tile scaling to make the effect visible
                    textureFill.TilePicOption = new TilePicOption
                    {
                        ScaleX = 0.5,
                        ScaleY = 0.5,
                        OffsetX = 0,
                        OffsetY = 0
                    };
                }
                else
                {
                    // Apply solid fill
                    shape.Fill.FillType = FillType.Solid;
                    shape.Fill.SolidFill.Color = Color.FromArgb(180, 200, 255); // Light blue
                }

                // Optional: give the shape a thin border for visual clarity
                shape.HasLine = true;
                shape.Line.Weight = 1;
                shape.Line.SolidFill.Color = Color.DarkBlue;
            }

            return workbook;
        }

        /// <param name="workbook">Workbook to render.</param>
        /// <param name="outputFileName">File name for the generated image.</param>
        /// <returns>TimeSpan representing the rendering duration.</returns>
        private static TimeSpan RenderWorksheetToImage(Workbook workbook, string outputFileName)
        {
            try
            {
                var options = new ImageOrPrintOptions
                {
                    // ImageFormat property is optional; default is PNG.
                    OnePagePerSheet = true,
                    Transparent = false,
                    HorizontalResolution = 96,
                    VerticalResolution = 96
                };

                var sheet = workbook.Worksheets[0];
                var renderer = new SheetRender(sheet, options);

                var sw = Stopwatch.StartNew();
                // Render the first (and only) page of the sheet
                renderer.ToImage(0, outputFileName);
                sw.Stop();

                return sw.Elapsed;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error rendering worksheet to image: {ex.Message}");
                return TimeSpan.Zero;
            }
        }
    }
}
