// Title: Render an Aspose.Cells worksheet to SVG and display it in a WPF Image control
// Description: Shows how to build a Workbook, export the first worksheet as an SVG file with Aspose.Cells SheetRender, and load the SVG into a WPF Image control for a desktop preview.
// Keywords: Aspose.Cells | SVG rendering | SheetRender | WPF Image control | C# | .NET | FitToViewPort | Excel to SVG | temporary SVG file | desktop preview
// Common Searches: Aspose.Cells export worksheet to SVG C# | Load SVG into WPF Image control | SheetRender SVG example | Display Excel sheet as SVG in WPF | C# render Excel as SVG for desktop app
// Developer Intent: Generate an SVG representation of an Excel worksheet and bind it to a WPF Image control for on‑screen display.
// Use Cases: Quick visual preview of spreadsheet data in a Windows desktop application without opening Excel. | Create a lightweight, scalable image of a sheet for reporting dashboards built with WPF. | Render multiple worksheets to SVG on the fly and swap the Image.Source at runtime based on user selection. | Handle rendering failures by checking the SVG file existence and falling back to a placeholder image.
// AI Prompts: Write C# code that loads the SVG file produced by SheetRender into a WPF Image control using SvgImageSource or a compatible library. | Show XAML and view‑model binding to display a temporary SVG path returned by RenderWorksheetToSvg() in a WPF Image control. | Provide error‑handling logic for missing or corrupted SVG files when updating the Image.Source in a WPF application.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeSvgDemo
{
    // Shows how to build a Workbook, export the first worksheet as an SVG file with Aspose.Cells SheetRender, and load the SVG into a WPF Image control for a desktop preview.
    class Program
    {
        [STAThread]
        static void Main()
        {
            try
            {
                string svgPath = RenderWorksheetToSvg();
                Console.WriteLine($"SVG file generated at: {svgPath}");
                // Optionally open the SVG with the default viewer:
                // System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(svgPath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        private static string RenderWorksheetToSvg()
        {
            // 1. Create a workbook and populate it with sample data
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(85);

            // 2. Configure SVG rendering options
            var svgOptions = new SvgImageOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Svg,
                FitToViewPort = true
            };

            // 3. Render the worksheet to a temporary SVG file
            var renderer = new SheetRender(sheet, svgOptions);
            string tempSvgPath = Path.Combine(Path.GetTempPath(), $"worksheet_{Guid.NewGuid()}.svg");
            renderer.ToImage(0, tempSvgPath);

            // 4. Verify the file was created
            if (!File.Exists(tempSvgPath))
                throw new FileNotFoundException("SVG file was not generated.", tempSvgPath);

            return tempSvgPath;
        }
    }
}
