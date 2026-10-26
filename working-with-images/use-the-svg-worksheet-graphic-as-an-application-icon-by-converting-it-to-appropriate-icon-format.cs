// Title: Convert an Excel worksheet to a PNG image using Aspose.Cells for .NET (C#)
// Common Searches: C# code to save an Excel worksheet as a PNG file with Aspose.Cells | Aspose.Cells .NET example for exporting a worksheet to PNG | How to generate a PNG image from an Excel workbook in C# | Convert Excel sheet to PNG using Aspose.Cells SaveFormat.Png | Save first worksheet of a workbook as PNG image in C# Aspose.Cells
// Tags: Aspose.Cells worksheet to PNG conversion | C# export Excel sheet as PNG image | SaveFormat.Png usage in Aspose.Cells | Excel workbook image generation with Aspose.Cells | Aspose.Cells image export example C#

using System;
using System.IO;
using Aspose.Cells;

// The program checks for 'input.xlsx', loads it with Aspose.Cells, and saves the first worksheet as a PNG file named 'output.png', handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputImagePath = "output.png";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook that contains the graphic
            Workbook workbook = new Workbook(inputPath);

            // Save the first worksheet (or the whole workbook) as a PNG image
            workbook.Save(outputImagePath, SaveFormat.Png);

            Console.WriteLine($"Image successfully saved to '{outputImagePath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
