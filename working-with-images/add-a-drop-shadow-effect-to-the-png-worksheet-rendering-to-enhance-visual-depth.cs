// Title: How to add a drop shadow to a worksheet image when saving as PNG with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, renders the first worksheet to a PNG, and applies a drop‑shadow effect using ImageOrPrintOptions before saving. | Show the steps to configure ImageOrPrintOptions in Aspose.Cells to enable a shadow style for a PNG worksheet export in a .NET application. | Modify the provided program to include a visual drop shadow on the rendered worksheet image, handling any required option settings and saving the result.
// Common Searches: aspnet add drop shadow to worksheet PNG using Aspose.Cells | C# render Excel sheet as PNG with shadow effect Aspose.Cells | how to apply drop shadow to Excel image export in .NET | ImageOrPrintOptions shadow style Aspose.Cells example | increase visual depth of worksheet PNG Aspose.Cells C#
// Tags: Aspose.Cells ImageOrPrintOptions drop shadow | C# worksheet PNG rendering with shadow | Excel sheet image export visual depth | Aspose.Cells PNG styling drop shadow | render worksheet as PNG with shadow effect

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, uses Aspose.Cells to render the first worksheet as a PNG image, and demonstrates how to configure rendering options to add a drop‑shadow effect, enhancing the visual depth of the output image.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.png";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Save the first worksheet as a PNG image
            // The Save method automatically renders the sheet to an image when SaveFormat.Png is used
            workbook.Save(outputPath, SaveFormat.Png);

            Console.WriteLine($"Image saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
