using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class WorksheetToImage
{
    static void Main()
    {
        // Load the source Excel file (XLSX)
        string sourcePath = "input.xlsx";
        Workbook workbook = new Workbook(sourcePath);

        // Access the first worksheet in the workbook
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure image rendering options
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = Aspose.Cells.Drawing.ImageType.Png; // Output format
        options.OnePagePerSheet = true; // Render the whole sheet on a single page

        // Create a SheetRender instance for the worksheet
        SheetRender sheetRender = new SheetRender(worksheet, options);

        // Define the output image file path
        string outputImagePath = "worksheet_image.png";

        // Render the first (and only) page of the sheet to the image file
        sheetRender.ToImage(0, outputImagePath);

        Console.WriteLine($"Worksheet successfully rendered to image: {outputImagePath}");
    }
}