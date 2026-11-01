// Title: Create an animated GIF from all worksheets in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to export each worksheet as a frame in an animated GIF with Aspose.Cells C# | Aspose.Cells save workbook as animated GIF each sheet becomes a frame | C# generate animated GIF from multiple Excel worksheets using Aspose.Cells | convert Excel workbook to animated GIF using Aspose.Cells .NET API | sample code to create GIF animation from Excel sheets in C#
// Tags: export workbook to animated GIF Aspose.Cells | worksheet to GIF frame Aspose.Cells | save Excel as GIF animation C# | Aspose.Cells multi-sheet GIF creation | C# generate animated GIF from Excel

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, verifies its presence, and uses Aspose.Cells to save the entire workbook as an animated GIF where each worksheet is rendered as a separate frame, with error handling for robustness.
class AnimatedGifCreator
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.gif";

        try
        {
            // Verify that the input workbook exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook as an animated GIF. Each worksheet becomes a frame.
            workbook.Save(outputPath, SaveFormat.Gif);

            Console.WriteLine($"Animated GIF created successfully at: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
