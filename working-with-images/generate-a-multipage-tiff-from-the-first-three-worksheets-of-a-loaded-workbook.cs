using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.tiff";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Determine which worksheets to render (up to the first three)
            int sheetCount = Math.Min(3, workbook.Worksheets.Count);
            int[] sheetIndexes = new int[sheetCount];
            for (int i = 0; i < sheetCount; i++)
                sheetIndexes[i] = i;

            // Configure rendering options for a multi‑page TIFF
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = ImageType.Tiff,
                SheetSet = new SheetSet(sheetIndexes)
            };

            // Render the selected sheets to a TIFF file
            WorkbookRender renderer = new WorkbookRender(workbook, options);
            renderer.ToImage(outputPath);

            Console.WriteLine("Multi‑page TIFF created successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}