// Title: C# – Add Drop Shadow to a Shape and Render a Worksheet as PNG with Aspose.Cells
// Description: Demonstrates how to create a workbook, insert data, add a rectangle shape, configure its ShadowEffect (preset type, color, transparency, blur, distance, angle), and export the worksheet to a PNG image using ImageOrPrintOptions while also saving the workbook with the shadowed shape.
// Keywords: Aspose.Cells C# drop shadow | render worksheet to PNG | shape ShadowEffect Aspose.Cells | ImageOrPrintOptions PNG export | preset shadow type OffsetBottom | Excel shape shadow C# | Aspose.Cells image rendering | add rectangle shape Aspose.Cells
// Common Searches: how to apply a drop shadow to a shape in Aspose.Cells | C# render Excel sheet as PNG with shadowed rectangle | Aspose.Cells shadow transparency blur distance | export worksheet to PNG with visual effects | Aspose.Cells ImageOrPrintOptions shadow example
// Developer Intent: The developer wants to apply a drop‑shadow effect to a shape and generate a PNG image of the worksheet that preserves the visual styling.
// Use Cases: Create polished reports where highlighted sections are emphasized with a shadowed rectangle and shared as PNG images on websites or dashboards. | Generate thumbnail previews of Excel sheets for UI galleries, adding depth with drop shadows to improve visual appeal. | Produce presentation‑ready workbooks that retain shadowed shapes while also providing PNG exports for non‑editable distribution.
// AI Prompts: Show me how to change the shadow color, blur radius, and distance for a shape in Aspose.Cells before exporting to PNG. | Provide C# code that applies different PresetShadowType values to multiple shapes and saves each worksheet as a separate PNG file. | Explain how to customize ShadowEffect angle and transparency for dynamic visual effects in Aspose.Cells rendering.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsDropShadowDemo
{
    // Demonstrates how to create a workbook, insert data, add a rectangle shape, configure its ShadowEffect (preset type, color, transparency, blur, distance, angle), and export the worksheet to a PNG image using ImageOrPrintOptions while also saving the workbook with the shadowed shape.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a rectangle shape that will carry the drop shadow
            // Parameters: upper left row, upper left row offset, upper left column, upper left column offset, height, width
            Shape shape = sheet.Shapes.AddRectangle(5, 0, 1, 0, 100, 200);

            // Configure the shadow effect (drop shadow)
            ShadowEffect shadow = shape.ShadowEffect;
            shadow.PresetType = PresetShadowType.OffsetBottom; // preset drop shadow
            shadow.Color = workbook.CreateCellsColor();        // create a color object
            shadow.Color.Color = System.Drawing.Color.Gray;    // set shadow color
            shadow.Transparency = 0.4;                         // 40% transparent
            shadow.Blur = 20;                                  // moderate blur
            shadow.Distance = 10;                              // distance from shape
            shadow.Angle = 90;                                 // shadow direction (downwards)

            // Set image rendering options for PNG output
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png,
                OnePagePerSheet = true
            };

            // Create SheetRender and render the first page to a PNG file
            SheetRender renderer = new SheetRender(sheet, options);
            renderer.ToImage(0, "RenderedWithShadow.png");

            // Optionally save the workbook to see the shape with shadow in Excel
            workbook.Save("WorkbookWithShadow.xlsx");

            Console.WriteLine("Worksheet rendered to PNG with drop shadow effect.");
        }
    }
}
