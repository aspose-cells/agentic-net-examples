// Title: Add a drop shadow to a shape and render the worksheet as a PNG with Aspose.Cells for .NET
// Common Searches: C# how to apply a drop shadow to a shape before exporting a worksheet to PNG using Aspose.Cells | Aspose.Cells render worksheet to PNG with rectangle shape shadow effect | set shadow blur, distance and transparency for a shape in Aspose.Cells PNG export | use ImageOrPrintOptions to generate high‑resolution PNG of a worksheet with visual depth
// Tags: add drop shadow to shape Aspose.Cells | worksheet to PNG rendering Aspose.Cells | configure ShadowEffect properties C# | ImageOrPrintOptions PNG resolution Aspose.Cells | preset shadow OffsetBottom Aspose.Cells | export worksheet with visual depth Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace WorksheetRenderWithShadow
{
    // The example creates a workbook, fills cells with data, adds a rectangle shape, configures its ShadowEffect (preset type, color, blur, distance, angle, transparency, size), sets PNG rendering options, renders the first worksheet page to a high‑resolution PNG file, and saves the workbook for verification.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and obtain the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Populate sample data.
                Cells cells = sheet.Cells;
                for (int row = 0; row < 10; row++)
                {
                    for (int col = 0; col < 5; col++)
                    {
                        cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // 3. Add a rectangle shape that will carry the drop shadow.
                Shape rect = sheet.Shapes.AddRectangle(
                    2,    // upperLeftRow
                    1,    // upperLeftColumn
                    0,    // upperLeftRowOffset (pixels)
                    0,    // upperLeftColumnOffset (pixels)
                    200,  // width (pixels)
                    100   // height (pixels)
                );

                // Set a solid fill type (default color will be used).
                rect.Fill.FillType = FillType.Solid;
                // Note: ForeColor property may not be available in some versions; omitted for compatibility.

                // 4. Configure the drop shadow effect.
                ShadowEffect shadow = rect.ShadowEffect;
                shadow.PresetType = PresetShadowType.OffsetBottom;
                shadow.Color = workbook.CreateCellsColor();
                shadow.Color.Color = Color.FromArgb(128, 0, 0, 0); // semi‑transparent black
                shadow.Blur = 20;
                shadow.Distance = 10;
                shadow.Angle = 45;
                shadow.Transparency = 0.4;
                shadow.Size = 1.0;

                // 5. Prepare rendering options for PNG output.
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    OnePagePerSheet = true,
                    HorizontalResolution = 200,
                    VerticalResolution = 200
                };

                // 6. Render the worksheet (first page) to a PNG file.
                SheetRender renderer = new SheetRender(sheet, renderOptions);
                string outputPath = "WorksheetWithShadow.png";

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                if (renderer.PageCount > 0)
                {
                    renderer.ToImage(0, outputPath);
                    Console.WriteLine($"Worksheet rendered successfully to '{outputPath}'.");
                }
                else
                {
                    Console.WriteLine("No pages were generated for the worksheet.");
                }

                // 7. Save the workbook for verification.
                string workbookPath = "WorksheetWithShadow.xlsx";
                workbook.Save(workbookPath);
                Console.WriteLine($"Workbook saved as '{workbookPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
