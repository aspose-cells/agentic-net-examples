using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Rendering; // For SvgImageOptions
using Aspose.Cells.Drawing;   // For ImageType

class SaveWorksheetAsSvgUtf8
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate cells with Unicode characters to demonstrate UTF‑8 preservation
        sheet.Cells["A1"].PutValue("中文字符");   // Chinese
        sheet.Cells["A2"].PutValue("😀 Emoji"); // Emoji
        sheet.Cells["A3"].PutValue("Привет");   // Cyrillic

        // Configure SVG rendering options
        SvgImageOptions svgOptions = new SvgImageOptions
        {
            // Ensure the output format is SVG
            ImageType = ImageType.Svg,

            // Fit the generated SVG to the viewport (optional, but commonly used)
            FitToViewPort = true
        };

        // Render the worksheet to an SVG file.
        // The ToImage method writes the SVG content using UTF‑8 encoding internally.
        SheetRender renderer = new SheetRender(sheet, svgOptions);
        renderer.ToImage(0, "WorksheetOutput.svg"); // Saves the SVG file

        // Optional: Save the workbook itself in XLSX format (demonstrates use of the save rule)
        workbook.Save("WorkbookOutput.xlsx", SaveFormat.Xlsx);

        Console.WriteLine("Worksheet saved as SVG with UTF‑8 encoding: WorksheetOutput.svg");
    }
}