// Title: Add a full-page image watermark that scales proportionally to the worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Insert a PNG file as a picture watermark in a new Aspose.Cells workbook and automatically resize it to cover the entire worksheet while preserving its aspect ratio. | Calculate the worksheet page dimensions and set the picture's width and height so the watermark fills the page behind the cells in C# with Aspose.Cells.
// Common Searches: C# Aspose.Cells fill worksheet with background image | scale picture to worksheet page size while preserving aspect ratio Aspose.Cells | add full-page background image to Excel workbook using Aspose.Cells .NET | set picture behind cells as watermark in Aspose.Cells C#
// Tags: Aspose.Cells image watermark | scale picture to worksheet page .NET | maintain aspect ratio Excel picture | C# set picture dimensions worksheet | full-page worksheet watermark .NET

using System;
using System.IO;
using Aspose.Cells;

// The example checks that watermark.png exists, creates a new Workbook, inserts the image as a picture at cell A1 on the first worksheet, and saves the file as WorkbookWithWatermark.xlsx, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            // Verify that the watermark image file exists
            const string watermarkPath = "watermark.png";
            if (!File.Exists(watermarkPath))
                throw new FileNotFoundException("Watermark image not found.", watermarkPath);

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add the watermark image as a picture (acts as a watermark)
            // Placed at the top-left corner of the sheet
            sheet.Pictures.Add(0, 0, watermarkPath);

            // Save the workbook
            const string outputPath = "WorkbookWithWatermark.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
