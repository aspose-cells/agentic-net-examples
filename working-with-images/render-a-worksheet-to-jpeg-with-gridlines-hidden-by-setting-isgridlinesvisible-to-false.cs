// Title: How to export an Excel worksheet to JPEG without gridlines using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export worksheet to JPEG with gridlines turned off | hide Excel gridlines when rendering to JPEG using Aspose.Cells .NET | Aspose.Cells render sheet as JPEG image without showing gridlines | set IsGridLinesVisible false before saving worksheet as JPEG in C# | ImageOrPrintOptions JPEG export hide gridlines Aspose.Cells example
// Tags: Aspose.Cells worksheet to JPEG export | gridlines visibility Aspose.Cells ImageOrPrintOptions | Excel to JPEG rendering with Aspose.Cells | C# SheetRender JPEG output | disable gridlines Aspose.Cells rendering

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a Workbook, adds sample data, optionally hides gridlines by setting worksheet.IsGridLinesVisible = false, configures ImageOrPrintOptions (150 dpi), uses SheetRender to render the first page, and saves the result as a JPEG file named 'WorksheetOutput.jpg'.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Optional: add some sample data to the worksheet
            worksheet.Cells["A1"].PutValue("Sample Text");
            worksheet.Cells["B2"].PutValue(12345);

            // Set image rendering options (default format is PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                HorizontalResolution = 150,
                VerticalResolution = 150
            };

            // Render the worksheet to an image
            SheetRender sheetRender = new SheetRender(worksheet, imgOptions);

            // Save the rendered image (page index 0) as a JPEG file
            string outputPath = "WorksheetOutput.jpg";
            sheetRender.ToImage(0, outputPath);
            Console.WriteLine($"Worksheet rendered and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
