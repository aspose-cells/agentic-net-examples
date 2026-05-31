using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class ExportWorksheetToImage
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        string workbookPath = "input.xlsx";
        Workbook workbook = new Workbook(workbookPath);

        // Select the worksheet to export (by index or name)
        int sheetIndex = 0; // first worksheet
        Worksheet sheet = workbook.Worksheets[sheetIndex];

        // Set up image rendering options
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Png;          // PNG supports transparency
        options.HorizontalResolution = 300;         // 300 DPI horizontal
        options.VerticalResolution = 300;           // 300 DPI vertical
        options.Transparent = true;                 // make background transparent
        options.OnePagePerSheet = true;             // render the whole sheet on one page

        // Create a SheetRender instance for the selected worksheet
        SheetRender renderer = new SheetRender(sheet, options);

        // Define the output image file path
        string imagePath = "output_sheet.png";

        // Render the first (and only) page of the worksheet to the image file
        renderer.ToImage(0, imagePath);

        Console.WriteLine($"Worksheet exported successfully to: {imagePath}");
    }
}