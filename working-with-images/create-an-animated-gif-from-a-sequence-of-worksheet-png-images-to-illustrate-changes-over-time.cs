// Title: Create a GIF image from an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to export the first worksheet of an Excel file to a GIF using Aspose.Cells in C# | Aspose.Cells C# example for saving a workbook as a GIF image | convert Excel workbook to GIF format with Aspose.Cells .NET library | C# code to generate a GIF file from an Excel workbook using Aspose.Cells | using SaveFormat.Gif in Aspose.Cells to create a GIF from Excel data
// Tags: Aspose.Cells save workbook as GIF | C# export Excel worksheet to GIF | Aspose.Cells SaveFormat.Gif example | convert Excel to GIF image .NET | generate GIF from Excel data using Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// This C# example loads an Excel file with Aspose.Cells, checks for its existence, and saves the first worksheet as a GIF image using the SaveFormat.Gif option, handling errors and confirming successful creation.
public class AnimatedGifCreator
{
    public static void Main()
    {
        try
        {
            // Path to the source Excel file.
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the workbook that contains the worksheets to be visualized.
            Workbook workbook = new Workbook(inputPath);

            // Path for the resulting GIF file.
            const string outputGifPath = "animated.gif";

            // Save the workbook as a GIF image.
            // Aspose.Cells will render the first worksheet to the GIF.
            // If you need multiple frames, additional processing would be required.
            workbook.Save(outputGifPath, SaveFormat.Gif);

            Console.WriteLine($"GIF file created successfully at '{outputGifPath}'.");
        }
        catch (Exception ex)
        {
            // Log or handle exceptions as needed.
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
