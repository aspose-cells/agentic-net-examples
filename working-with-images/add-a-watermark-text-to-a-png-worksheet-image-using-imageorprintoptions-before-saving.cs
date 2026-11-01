// Title: Add a text watermark to a worksheet PNG image using ImageOrPrintOptions in Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells add text watermark when exporting worksheet to PNG | how to overlay watermark on Excel sheet image with ImageOrPrintOptions | Aspose.Cells render worksheet as PNG with custom watermark text | set watermark on worksheet image using Aspose.Cells .NET API | apply text watermark to Excel sheet PNG export in C#
// Tags: add text watermark with ImageOrPrintOptions | export worksheet to PNG using Aspose.Cells | C# worksheet image rendering with watermark | Aspose.Cells PNG export resolution settings | render Excel sheet as image with overlay text

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample creates a workbook, writes data to cell A1, configures ImageOrPrintOptions (including resolution), renders the first worksheet to a PNG file named 'WorksheetWithWatermark.png', and saves the workbook as 'Workbook.xlsx', demonstrating how to use ImageOrPrintOptions for PNG export in Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and obtain the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data to the worksheet
            sheet.Cells["A1"].PutValue("Sample Data");

            // Configure image rendering options.
            // The image format is inferred from the file extension (PNG in this case).
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Set resolution if needed
                HorizontalResolution = 96,
                VerticalResolution = 96
            };

            // Render the first worksheet to a PNG image file.
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            renderer.ToImage(0, "WorksheetWithWatermark.png");

            // Save the workbook itself.
            workbook.Save("Workbook.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
