// Title: Export an Aspose.Cells worksheet to a JPEG file while suppressing gridlines using C#
// AI Prompts: Write C# code that uses Aspose.Cells to save a worksheet as a JPEG image with the gridlines turned off. | Demonstrate how to configure PageSetup.PrintGridlines and ImageOrPrintOptions to generate a JPEG without visible gridlines. | Convert a specific worksheet page to JPEG in C# while ensuring the rendered image does not display Excel gridlines.
// Common Searches: Aspose.Cells C# export worksheet to JPEG without showing gridlines | How to hide Excel gridlines when rendering to image with Aspose.Cells | C# code sample for saving Excel sheet as JPEG image with PrintGridlines false | Render only the first page of a workbook to JPEG using Aspose.Cells and remove gridlines | ImageOrPrintOptions settings to disable gridlines in JPEG output with Aspose.Cells
// Tags: Aspose.Cells worksheet to JPEG conversion | disable gridlines in Aspose.Cells image rendering | PageSetup.PrintGridlines false C# | SheetRender ToImage without gridlines | ImageOrPrintOptions JPEG output Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// // Creates a workbook, disables gridline printing via PageSetup.PrintGridlines, and uses SheetRender with default ImageOrPrintOptions to save the first worksheet page as a JPEG file without gridlines.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Hide gridlines when printing/rendering
            sheet.PageSetup.PrintGridlines = false;

            // (Optional) Add some sample data
            sheet.Cells["A1"].PutValue("Sample Text");
            sheet.Cells["B2"].PutValue(12345);

            // Set up rendering options (default options are sufficient for JPEG output)
            ImageOrPrintOptions options = new ImageOrPrintOptions();

            // Render the worksheet to an image
            SheetRender renderer = new SheetRender(sheet, options);

            // Save the first page (the worksheet) as a JPEG file
            renderer.ToImage(0, "WorksheetOutput.jpg");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
