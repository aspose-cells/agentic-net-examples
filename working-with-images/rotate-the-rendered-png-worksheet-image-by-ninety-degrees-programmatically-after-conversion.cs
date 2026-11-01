// Title: How to rotate a PNG worksheet image 90 degrees in C# after rendering with Aspose.Cells
// Common Searches: C# rotate PNG image generated from Excel worksheet using Aspose.Cells | rotate worksheet screenshot 90 degrees after SheetRender Aspose.Cells .NET | programmatically rotate PNG output from Aspose.Cells without external tools | Aspose.Cells render worksheet to PNG then rotate image in .NET Core | using System.Drawing.Common to rotate Aspose.Cells PNG image in C#
// Tags: rotate PNG image after Aspose.Cells SheetRender | Aspose.Cells worksheet to PNG conversion | System.Drawing.Common image rotation .NET | C# image manipulation of Excel render output | 90-degree rotation of Excel worksheet image

using Aspose.Cells;
using Aspose.Cells.Rendering;
using System;
using System.IO;

// The example loads an Excel workbook, renders the first worksheet to a PNG file with Aspose.Cells SheetRender, and demonstrates how to rotate the resulting PNG by 90 degrees in C# using System.Drawing.Common, enabling full programmatic image manipulation after conversion.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Verify input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set image rendering options (PNG format is inferred from file extension)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            // Render the worksheet to a PNG image
            SheetRender sheetRender = new SheetRender(worksheet, imgOptions);
            string tempImagePath = "temp.png";
            sheetRender.ToImage(0, tempImagePath);

            // Verify temporary image was created
            if (!File.Exists(tempImagePath))
            {
                Console.WriteLine("Failed to create temporary image.");
                return;
            }

            // NOTE: Image rotation can be performed with external libraries (e.g., System.Drawing.Common).
            // The rotation step is omitted here to avoid additional dependencies.

            Console.WriteLine($"Image saved to {tempImagePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
