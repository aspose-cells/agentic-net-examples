// Title: Export the first textbox shape from an Excel worksheet to a transparent PNG using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, finds the first textbox shape on the first sheet, sets its fill transparency to 100 % and saves the sheet as a PNG image with Aspose.Cells. | Show how to configure ImageOrPrintOptions for PNG output and apply shape.Fill.Transparency before rendering a worksheet to a transparent PNG in C#. | Provide a robust C# example that validates the workbook file, checks for a textbox shape, applies full transparency, and handles exceptions while exporting to PNG with Aspose.Cells.
// Common Searches: how to export a textbox shape from Excel to a PNG with transparent background using Aspose.Cells C# | Aspose.Cells set shape fill transparency before rendering worksheet to image | C# render only a specific shape from an Excel sheet as PNG with Aspose.Cells | save Excel textbox as transparent PNG programmatically Aspose.Cells .NET
// Tags: textbox shape export to PNG Aspose.Cells | shape fill transparency Aspose.Cells .NET | ImageOrPrintOptions PNG rendering Aspose.Cells | worksheet shape rendering transparent background C# | exception handling missing shape Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, locates the first shape (assumed to be a textbox) on the first worksheet, sets its fill transparency to fully transparent, and renders the worksheet to a PNG file using Aspose.Cells, with error handling for missing files or shapes.
class ExportTextboxAsPng
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "textbox.png";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the first shape (assumed to be a textbox)
            Shape textbox = null;
            foreach (Shape shape in sheet.Shapes)
            {
                textbox = shape;
                break;
            }

            if (textbox == null)
            {
                Console.WriteLine("No shape found on the worksheet.");
                return;
            }

            // Make the textbox background fully transparent (0 = opaque, 1 = fully transparent)
            textbox.Fill.Transparency = 1.0;

            // Set rendering options (default format is PNG)
            ImageOrPrintOptions options = new ImageOrPrintOptions();

            // Render the worksheet (including the textbox) to an image file
            SheetRender renderer = new SheetRender(sheet, options);
            renderer.ToImage(0, outputPath);

            Console.WriteLine($"Textbox exported successfully as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
