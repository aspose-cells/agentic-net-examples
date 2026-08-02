// Title: Add a Rendered Worksheet SVG as an Excel Icon with Aspose.Cells for .NET
// Description: Demonstrates how to render a worksheet to SVG in memory, embed the SVG as an Excel icon using ShapeCollection.AddIcons, set the picture to display as an icon, and save the workbook with the vector graphic as an embedded icon.
// Keywords: Aspose.Cells | C# | .NET | SVG icon | Excel workbook icon | ShapeCollection.AddIcons | render worksheet to SVG | embed SVG in Excel | display picture as icon | vector graphic Excel | code example
// Common Searches: embed SVG as icon in Excel using Aspose.Cells | Aspose.Cells render worksheet to SVG | C# add custom icon to Excel sheet | ShapeCollection AddIcons example | display picture as icon Aspose.Cells .NET
// Developer Intent: Embed a worksheet‑generated SVG as an Excel icon programmatically with Aspose.Cells for .NET.
// Use Cases: Create sales dashboards where each workbook shows a scalable SVG thumbnail as its icon for instant visual identification. | Automate branding of Excel templates by inserting vector‑based icons derived from sheet content. | Package multiple Excel reports with SVG icons that retain quality at any display size.
// AI Prompts: Generate C# code that converts a rendered worksheet SVG to a .ico file and sets it as the workbook’s file icon using Aspose.Cells. | Explain how to adjust the size and position of an SVG icon added with ShapeCollection.AddIcons in Aspose.Cells. | Show how to add different SVG icons to several worksheets and configure each to display as an icon.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsIconExample
{
    // Demonstrates how to render a worksheet to SVG in memory, embed the SVG as an Excel icon using ShapeCollection.AddIcons, set the picture to display as an icon, and save the workbook with the vector graphic as an embedded icon.
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

            // Render the worksheet to SVG and capture the SVG bytes in memory
            SvgImageOptions svgOptions = new SvgImageOptions
            {
                ImageType = ImageType.Svg,   // Ensure SVG output
                FitToViewPort = true
            };

            byte[] svgBytes;
            using (MemoryStream svgStream = new MemoryStream())
            {
                // Render the first (and only) page of the worksheet to SVG
                SheetRender renderer = new SheetRender(sheet, svgOptions);
                renderer.ToImage(0, svgStream);
                svgBytes = svgStream.ToArray();
            }

            // Add the SVG as an icon to the worksheet using ShapeCollection.AddIcons
            ShapeCollection shapes = sheet.Shapes;

            // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
            // Using -1 for height/width lets Excel auto‑size the icon.
            Picture iconPicture = shapes.AddIcons(
                topRow: 5,          // place starting at row 6 (0‑based index)
                top: 0,
                leftColumn: 2,      // place starting at column C
                left: 0,
                height: -1,
                width: -1,
                imageByteData: svgBytes,
                compatibleImageData: null);

            // Mark the picture to be displayed as an icon
            iconPicture.DisplayAsIcon = true;
            iconPicture.Name = "WorksheetSvgIcon";

            // Save the workbook (the SVG icon is now embedded)
            workbook.Save("WorksheetWithSvgIcon.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved with SVG used as an application icon.");
        }
    }
}
