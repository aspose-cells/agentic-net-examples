// Title: Embed a Worksheet SVG as an Excel Icon with Aspose.Cells for .NET
// Description: Demonstrates how to render a worksheet to SVG, read the SVG bytes, and embed the image as an icon on the sheet using ShapeCollection.AddIcons. The example sets DisplayAsIcon, custom dimensions, and saves the workbook as XLSX.
// Keywords: Aspose.Cells SVG icon | AddIcons C# | DisplayAsIcon property | render worksheet to SVG | embed SVG in Excel | Aspose.Cells .NET example | Excel icon from SVG | C# Excel image embedding
// Common Searches: How to embed an SVG as an icon in Excel using Aspose.Cells | Aspose.Cells AddIcons method example | Set DisplayAsIcon for a picture in Aspose.Cells | Render worksheet to SVG and use as Excel thumbnail | C# embed scalable SVG in XLSX file
// Developer Intent: Add a rendered worksheet SVG as an embedded icon inside an Excel workbook with Aspose.Cells for .NET.
// Use Cases: Create workbooks that show a small SVG preview of each sheet for quick visual navigation. | Add a custom SVG logo as an icon on a template sheet to reinforce branding. | Automate report generation that includes a scalable SVG thumbnail of a chart as an icon.
// AI Prompts: Generate code to add a PNG fallback for older Excel versions when embedding an SVG icon with Aspose.Cells. | Show how to scale an SVG icon dynamically to fit a target cell range while keeping aspect ratio. | Explain how to place multiple different SVG icons on separate worksheets using ShapeCollection.AddIcons.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsIconDemo
{
    // Demonstrates how to render a worksheet to SVG, read the SVG bytes, and embed the image as an icon on the sheet using ShapeCollection.AddIcons. The example sets DisplayAsIcon, custom dimensions, and saves the workbook as XLSX.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data (optional, just to have content)
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(95);

            // -----------------------------------------------------------------
            // Render the worksheet to an SVG image (this will be used as the icon)
            // -----------------------------------------------------------------
            string svgPath = "worksheet_icon.svg";

            // Configure SVG rendering options
            SvgImageOptions svgOptions = new SvgImageOptions
            {
                ImageType = ImageType.Svg,   // Ensure SVG output
                FitToViewPort = true         // Fit the SVG to the viewport
            };

            // Render the first (and only) sheet to SVG
            SheetRender renderer = new SheetRender(sheet, svgOptions);
            renderer.ToImage(0, svgPath);

            // Read the generated SVG bytes
            byte[] svgBytes = File.ReadAllBytes(svgPath);

            // ---------------------------------------------------------------
            // Add the SVG as an icon to the worksheet using ShapeCollection.AddIcons
            // ---------------------------------------------------------------
            ShapeCollection shapes = sheet.Shapes;

            // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset),
            // height, width, imageByteData (SVG), compatibleImageData (null for newer Excel)
            // Using -1 for height and width lets Excel auto-size the icon.
            Picture iconPicture = shapes.AddIcons(
                topRow: 5,          // place starting at row 6 (0‑based index)
                top: 0,
                leftColumn: 2,      // place starting at column C
                left: 0,
                height: -1,
                width: -1,
                imageByteData: svgBytes,
                compatibleImageData: null);

            // Mark the picture to be displayed as an icon (prevents auto‑conversion)
            iconPicture.DisplayAsIcon = true;

            // Optionally set a name and size for the icon picture
            iconPicture.Name = "WorksheetSvgIcon";
            iconPicture.Height = 64;   // pixels
            iconPicture.Width = 64;    // pixels

            // ---------------------------------------------------------------
            // Save the workbook (the SVG icon is now embedded)
            // ---------------------------------------------------------------
            string outputPath = "WorkbookWithSvgIcon.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            // Clean up temporary SVG file
            if (File.Exists(svgPath))
                File.Delete(svgPath);

            Console.WriteLine($"Workbook saved to '{outputPath}' with SVG icon embedded.");
        }
    }
}
