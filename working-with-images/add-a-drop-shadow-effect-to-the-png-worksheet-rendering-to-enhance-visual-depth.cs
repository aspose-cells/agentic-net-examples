// Title: C# – Add Drop Shadow to a Shape and Render Worksheet as PNG with Aspose.Cells
// Description: Demonstrates how to create a workbook, insert data, add a rectangle shape, apply a customizable drop‑shadow effect, and export the worksheet to a PNG image using Aspose.Cells for .NET. The workbook is also saved to preserve the shape for further Excel editing.
// Keywords: Aspose.Cells drop shadow C# | render worksheet to PNG Aspose.Cells | C# shape shadow effect | ImageOrPrintOptions PNG Aspose.Cells | SheetRender PNG export | Aspose.Cells shape styling | C# Excel image rendering
// Common Searches: how to apply drop shadow to a shape in Aspose.Cells | export Excel worksheet as PNG with shadow using C# | Aspose.Cells set shadow properties before image rendering | C# render worksheet to PNG with shape effects | Aspose.Cells PNG output with visual depth
// Developer Intent: Apply a drop‑shadow to a rectangle shape and generate a PNG image of the worksheet using Aspose.Cells for .NET.
// Use Cases: Create polished dashboard screenshots where shapes have depth. | Produce PNG reports that retain visual styling from Excel. | Save an Excel file with shadowed shapes for later editing while providing a ready‑to‑use PNG preview.
// AI Prompts: Show C# code to modify shadow angle, blur, and transparency for a shape before PNG export with Aspose.Cells. | Generate an example that adds multiple shapes, each with a different preset shadow, and saves each sheet page as a separate PNG. | Explain how to configure ImageOrPrintOptions to keep shape shadows when converting a worksheet to high‑resolution PNG.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsDropShadowDemo
{
    // Demonstrates how to create a workbook, insert data, add a rectangle shape, apply a customizable drop‑shadow effect, and export the worksheet to a PNG image using Aspose.Cells for .NET. The workbook is also saved to preserve the shape for further Excel editing.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(85);
            sheet.Cells["A4"].PutValue("Cherries");
            sheet.Cells["B4"].PutValue(60);

            // Add a rectangle shape that will act as a visual container
            // Parameters: upper left row, upper left row offset, upper left column, upper left column offset, height, width
            Shape rect = sheet.Shapes.AddRectangle(0, 0, 0, 0, 200, 300);

            // Configure the drop shadow effect for the shape
            ShadowEffect shadow = rect.ShadowEffect;
            shadow.PresetType = PresetShadowType.OffsetBottom; // simple offset shadow
            shadow.Blur = 20;          // moderate blur
            shadow.Distance = 10;      // distance from shape
            shadow.Transparency = 0.3; // 30% transparent
            shadow.Angle = 135;        // direction of the light source
            shadow.Size = 1.0;         // default size

            // Set image rendering options for PNG output
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png,
                OnePagePerSheet = true
            };

            // Create SheetRender instance using the worksheet and options
            SheetRender renderer = new SheetRender(sheet, options);

            // Render the first (and only) page to a PNG file
            string outputPath = "RenderedWithShadow.png";
            renderer.ToImage(0, outputPath);

            // Optionally save the workbook to verify the shape and its shadow in Excel
            workbook.Save("WorkbookWithShadow.xlsx");

            Console.WriteLine($"Worksheet rendered to PNG with drop shadow: {outputPath}");
        }
    }
}
