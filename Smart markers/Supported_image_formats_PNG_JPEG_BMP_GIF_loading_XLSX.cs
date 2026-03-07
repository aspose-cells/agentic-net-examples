using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class ImageExportDemo
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string sourcePath = "input.xlsx";

        // Load the workbook from the XLSX file (lifecycle rule: load)
        Workbook workbook = new Workbook(sourcePath);

        // Create image rendering options; default ImageType is PNG
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

        // Iterate through each worksheet in the workbook
        for (int sheetIdx = 0; sheetIdx < workbook.Worksheets.Count; sheetIdx++)
        {
            Worksheet sheet = workbook.Worksheets[sheetIdx];

            // Create a SheetRender for the current worksheet (uses the rendering rule)
            SheetRender sheetRender = new SheetRender(sheet, imgOptions);

            // Render every page of the worksheet to the supported image formats
            for (int pageIdx = 0; pageIdx < sheetRender.PageCount; pageIdx++)
            {
                string baseName = $"Sheet{sheetIdx}_Page{pageIdx}";

                // ----- PNG (default) -----
                imgOptions.ImageType = ImageType.Png;
                sheetRender.ToImage(pageIdx, $"{baseName}.png");

                // ----- JPEG -----
                imgOptions.ImageType = ImageType.Jpeg;
                sheetRender.ToImage(pageIdx, $"{baseName}.jpeg");

                // ----- BMP -----
                imgOptions.ImageType = ImageType.Bmp;
                sheetRender.ToImage(pageIdx, $"{baseName}.bmp");

                // ----- GIF -----
                imgOptions.ImageType = ImageType.Gif;
                sheetRender.ToImage(pageIdx, $"{baseName}.gif");
            }
        }

        Console.WriteLine("All sheets have been exported to PNG, JPEG, BMP, and GIF images.");
    }
}