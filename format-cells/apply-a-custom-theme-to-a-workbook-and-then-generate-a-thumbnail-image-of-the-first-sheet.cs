// Title: C# – Apply a Custom Theme to an Aspose.Cells Workbook and Export the First Sheet as a PNG Thumbnail
// Description: Creates a new Workbook, defines twelve custom theme colors, applies them with Workbook.CustomTheme, saves the file, configures ImageOrPrintOptions for a low‑resolution PNG, and uses WorkbookRender to generate a thumbnail of the first worksheet.
// Keywords: Aspose.Cells | C# | CustomTheme | Excel theme colors | WorkbookRender | PNG thumbnail | sheet preview image | ImageOrPrintOptions | low resolution export | .NET Excel rendering
// Common Searches: Aspose.Cells set custom theme C# | export first worksheet as PNG thumbnail Aspose.Cells | WorkbookRender generate sheet preview image | ImageOrPrintOptions low resolution thumbnail | save Excel file with custom colors Aspose
// Developer Intent: Apply a custom color theme to a workbook and create a PNG thumbnail of the first worksheet using Aspose.Cells for .NET.
// Use Cases: Brand corporate Excel reports with a company palette and provide a quick preview image for document portals. | Automatically generate low‑resolution sheet snapshots for file‑management systems that display workbook thumbnails. | Persist workbooks with a custom theme while supplying a PNG preview for email attachments or UI galleries.
// AI Prompts: Show how to modify the custom theme colors after the workbook has been created with Aspose.Cells. | Provide code to create thumbnails for every worksheet in a workbook, allowing configurable resolution and image format. | Explain how to embed the generated PNG thumbnail into a PDF using Aspose.PDF.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Creates a new Workbook, defines twelve custom theme colors, applies them with Workbook.CustomTheme, saves the file, configures ImageOrPrintOptions for a low‑resolution PNG, and uses WorkbookRender to generate a thumbnail of the first worksheet.
class CustomThemeThumbnailDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data to visualize the theme effect
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

        // Save the workbook (optional, demonstrates persistence of the theme)
        workbook.Save("CustomThemeWorkbook.xlsx");

        // Prepare image rendering options for the thumbnail
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = Aspose.Cells.Drawing.ImageType.Png,
            // Render only the first page (first sheet) as a thumbnail
            OnePagePerSheet = true,
            // Reduce resolution for a smaller thumbnail if desired
            HorizontalResolution = 150,
            VerticalResolution = 150
        };

        // Create a renderer for the workbook
        WorkbookRender renderer = new WorkbookRender(workbook, imgOptions);

        // Render the first page (index 0) to a PNG file as a thumbnail
        renderer.ToImage(0, "FirstSheetThumbnail.png");

        // Clean up resources
        renderer.Dispose();
        workbook.Dispose();

        Console.WriteLine("Custom theme applied and thumbnail generated successfully.");
    }
}
