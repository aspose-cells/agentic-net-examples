// Title: Render an Aspose.Cells worksheet to SVG and show it as PNG in a WPF Image control (C#)
// Description: The example builds a Workbook, fills cells with sample data, saves the first worksheet as an SVG file (SaveFormat.Svg) and renders the same sheet to a PNG image using SheetRender. The PNG can then be loaded into a WPF Image control for runtime display in a desktop application.
// Keywords: Aspose.Cells | C# | WPF Image control | SVG export | PNG rendering | SheetRender | SaveFormat.Svg | ImageOrPrintOptions | Excel to image | desktop UI
// Common Searches: Aspose.Cells export worksheet to SVG C# | How to render Excel sheet as PNG for WPF | Display Aspose.Cells PNG in WPF Image element | Save Excel worksheet as SVG and PNG using Aspose.Cells | C# code to load PNG into WPF Image control
// Developer Intent: Generate vector SVG and raster PNG versions of a worksheet and use the PNG for visual presentation in a WPF desktop UI.
// Use Cases: Create a high‑quality SVG for printing or web publishing. | Produce a PNG snapshot that can be bound to a WPF Image control. | Integrate Excel sheet visuals into a .NET desktop application without requiring Excel installation.
// AI Prompts: Write C# code that loads the generated worksheet.png into a WPF Image control and binds it in XAML. | Show how to increase PNG resolution with ImageOrPrintOptions when rendering an Aspose.Cells worksheet. | Provide a method to convert the SVG output to an ImageSource for direct display in WPF without saving to disk.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example builds a Workbook, fills cells with sample data, saves the first worksheet as an SVG file (SaveFormat.Svg) and renders the same sheet to a PNG image using SheetRender. The PNG can then be loaded into a WPF Image control for runtime display in a desktop application.
public class Program
{
    public static void Main()
    {
        try
        {
            // -------------------------------------------------
            // 1. Create a sample workbook and populate data
            // -------------------------------------------------
            Workbook workbook = new Workbook();                     // create workbook
            Worksheet sheet = workbook.Worksheets[0];               // get first worksheet
            sheet.Cells["A1"].Value = "Product";
            sheet.Cells["B1"].Value = "Sales";
            sheet.Cells["A2"].Value = "Apple";
            sheet.Cells["B2"].Value = 120;
            sheet.Cells["A3"].Value = "Orange";
            sheet.Cells["B3"].Value = 150;

            // -------------------------------------------------
            // 2. Render the worksheet to an SVG file
            // -------------------------------------------------
            string svgPath = "worksheet.svg";
            // Use SaveOptions for SVG since ImageOrPrintOptions may not expose ImageFormat in some versions
            workbook.Save(svgPath, SaveFormat.Svg);
            Console.WriteLine($"SVG saved to: {Path.GetFullPath(svgPath)}");

            // -------------------------------------------------
            // 3. Render the same worksheet to PNG (for display)
            // -------------------------------------------------
            string pngPath = "worksheet.png";
            ImageOrPrintOptions pngOptions = new ImageOrPrintOptions(); // default format is PNG
            SheetRender pngRender = new SheetRender(sheet, pngOptions);
            pngRender.ToImage(0, pngPath); // save PNG to file
            Console.WriteLine($"PNG saved to: {Path.GetFullPath(pngPath)}");
        }
        catch (Exception ex)
        {
            // Runtime safety: log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
