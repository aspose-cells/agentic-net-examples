using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample HTML Export with higher DPI images");

        // Initialize HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Access the ImageOrPrintOptions through HtmlSaveOptions.ImageOptions
        ImageOrPrintOptions imageOptions = saveOptions.ImageOptions;

        // Set the desired DPI for images (both horizontal and vertical)
        imageOptions.HorizontalResolution = 150; // 150 DPI horizontally
        imageOptions.VerticalResolution = 150;   // 150 DPI vertically

        // Optional: export images as separate files (not Base64) for easier inspection
        saveOptions.ExportImagesAsBase64 = false;

        // Save the workbook as HTML with the configured image DPI
        workbook.Save("output.html", saveOptions);

        Console.WriteLine("HTML file saved with images at 150 DPI.");
    }
}