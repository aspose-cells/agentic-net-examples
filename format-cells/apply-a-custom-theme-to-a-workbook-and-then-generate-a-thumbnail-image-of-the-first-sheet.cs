// Title: C# – Apply a Custom Theme and Export the First Worksheet as a PNG Thumbnail with Aspose.Cells
// Description: This example creates a new Workbook, defines a 12‑color custom theme, applies it via Workbook.CustomTheme, saves the file, configures ImageOrPrintOptions for PNG output, and uses WorkbookRender to generate a 96 DPI thumbnail of the first sheet.
// Keywords: Aspose.Cells custom theme C# | Workbook.CustomTheme .NET | WorkbookRender PNG thumbnail | ImageOrPrintOptions export image | Excel workbook preview image | C# generate sheet thumbnail | Aspose.Cells theme colors
// Common Searches: how to set a custom theme in Aspose.Cells for .NET | export first worksheet as PNG using Aspose.Cells | create Excel thumbnail with custom colors C# | Aspose.Cells render sheet to image example | apply corporate color palette to workbook Aspose
// Developer Intent: Apply a custom color theme to a workbook and produce a PNG thumbnail of its first sheet.
// Use Cases: Produce branded preview images for Excel reports displayed in web portals. | Automate distribution of themed workbooks with a quick visual identifier for email notifications. | Generate low‑resolution snapshots for document management systems that need sheet previews.
// AI Prompts: Write C# code that defines a 12‑color custom theme, applies it to an Aspose.Cells workbook, and saves the workbook. | Show how to configure ImageOrPrintOptions for a 96 DPI PNG and use WorkbookRender to export the first worksheet as a thumbnail. | Explain the steps to combine Workbook.CustomTheme with WorkbookRender to create a themed workbook and a corresponding sheet preview image.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This example creates a new Workbook, defines a 12‑color custom theme, applies it via Workbook.CustomTheme, saves the file, configures ImageOrPrintOptions for PNG output, and uses WorkbookRender to generate a 96 DPI thumbnail of the first sheet.
class CustomThemeAndThumbnailDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data to demonstrate the custom theme
        sheet.Cells["A1"].PutValue("Custom Theme Demo");
        sheet.Cells["A2"].PutValue(123);
        sheet.Cells["A3"].PutValue(DateTime.Now);

        // Define 12 custom theme colors (Background1, Text1, Background2, Text2, Accent1‑Accent6, Hyperlink, Followed Hyperlink)
        Color[] customColors = new Color[]
        {
            Color.FromArgb(255, 255, 255), // Background1 - White
            Color.FromArgb(0, 0, 0),       // Text1 - Black
            Color.FromArgb(240, 240, 240), // Background2 - Light Gray
            Color.FromArgb(80, 80, 80),    // Text2 - Dark Gray
            Color.FromArgb(255, 0, 0),     // Accent1 - Red
            Color.FromArgb(0, 255, 0),     // Accent2 - Green
            Color.FromArgb(0, 0, 255),     // Accent3 - Blue
            Color.FromArgb(255, 165, 0),   // Accent4 - Orange
            Color.FromArgb(128, 0, 128),   // Accent5 - Purple
            Color.FromArgb(0, 255, 255),   // Accent6 - Cyan
            Color.FromArgb(0, 0, 255),     // Hyperlink - Blue
            Color.FromArgb(128, 0, 0)      // Followed Hyperlink - Maroon
        };

        // Apply the custom theme to the workbook
        workbook.CustomTheme("MyCustomTheme", customColors);

        // Save the workbook (optional, demonstrates persistence)
        workbook.Save("CustomThemeWorkbook.xlsx");

        // Prepare image rendering options for the thumbnail
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = Aspose.Cells.Drawing.ImageType.Png,
            OnePagePerSheet = true,          // Ensure each sheet is rendered as a single page
            HorizontalResolution = 96,       // Typical screen DPI
            VerticalResolution = 96
        };

        // Create a renderer for the workbook
        WorkbookRender renderer = new WorkbookRender(workbook, imgOptions);

        // Render the first sheet (page index 0) to a PNG thumbnail file
        string thumbnailPath = "FirstSheetThumbnail.png";
        renderer.ToImage(0, thumbnailPath);

        Console.WriteLine($"Custom theme applied and thumbnail generated at: {thumbnailPath}");
    }
}
