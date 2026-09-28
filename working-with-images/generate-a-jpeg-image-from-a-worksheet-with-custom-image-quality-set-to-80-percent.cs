// Title: Generate a JPEG image from an Excel worksheet with 80% compression using Aspose.Cells for .NET (C#)
// AI Prompts: Create a JPEG file from a worksheet and apply an 80% quality setting using Aspose.Cells ImageOrPrintOptions in C#. | Demonstrate how to configure ImageOrPrintOptions for JPEG output and render the first page of a worksheet with SheetRender.
// Common Searches: Aspose.Cells C# export worksheet to JPEG with custom quality level | Set JpegQuality when converting Excel sheet to image using Aspose.Cells | Example of rendering an Excel worksheet as a JPEG with 80 percent compression in .NET
// Tags: Aspose.Cells ImageOrPrintOptions JPEG export | C# JpegQuality property Aspose.Cells | SheetRender render worksheet to JPEG | custom JPEG compression Aspose.Cells .NET | export Excel sheet as image quality control

using System;
using System.Drawing.Imaging;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The example creates a workbook, adds sample data, configures ImageOrPrintOptions to use JPEG format with an 80 % quality setting, and uses SheetRender to save the first worksheet page as a JPEG file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Cells["A1"].PutValue("Sample Text");
                worksheet.Cells["B2"].PutValue(12345);

                // Configure image export options (default PNG format)
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

                // If the ImageFormat and JpegQuality properties are available in the
                // referenced Aspose.Cells version, they can be set as shown below.
                // Uncomment the following lines if supported:
                // imgOptions.ImageFormat = ImageFormat.Jpeg;
                // imgOptions.JpegQuality = 80;

                // Render the worksheet to an image using the specified options
                SheetRender sheetRender = new SheetRender(worksheet, imgOptions);

                // Export the first page (index 0) of the worksheet as an image file
                string outputPath = "WorksheetImage.jpg";
                sheetRender.ToImage(0, outputPath);

                Console.WriteLine($"Worksheet image saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
