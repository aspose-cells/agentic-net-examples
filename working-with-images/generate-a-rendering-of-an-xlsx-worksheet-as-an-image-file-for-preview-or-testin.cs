using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class WorksheetToImageDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(150);
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["B3"].PutValue(200);

        // Configure image rendering options
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = Aspose.Cells.Drawing.ImageType.Png; // output format
        options.OnePagePerSheet = true; // render the sheet as a single page

        // Create a SheetRender instance for the worksheet
        SheetRender renderer = new SheetRender(sheet, options);

        // Prepare output paths
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);
        string imagePath = Path.Combine(outputDir, "SheetPreview.png");
        string workbookPath = Path.Combine(outputDir, "SampleWorkbook.xlsx");

        // Render the first (and only) page of the sheet to an image file
        renderer.ToImage(0, imagePath);

        // Save the original workbook for reference
        workbook.Save(workbookPath);

        Console.WriteLine($"Worksheet rendered to image: {imagePath}");
        Console.WriteLine($"Workbook saved to: {workbookPath}");
    }
}