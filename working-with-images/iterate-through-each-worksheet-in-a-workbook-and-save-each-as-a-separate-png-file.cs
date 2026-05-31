using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Path to the source workbook
        string inputPath = "input.xlsx";

        // Directory where PNG files will be saved
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Load the workbook
        Workbook workbook = new Workbook(inputPath);

        // Configure rendering options for PNG output
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = Aspose.Cells.Drawing.ImageType.Png,
            OnePagePerSheet = true // Ensure each sheet renders to a single page
        };

        // Iterate through each worksheet and save as a separate PNG file
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet sheet = workbook.Worksheets[i];

            // Create a SheetRender for the current worksheet
            SheetRender sheetRender = new SheetRender(sheet, options);

            // Build a safe file name using the worksheet name
            string safeName = string.Concat(sheet.Name.Split(Path.GetInvalidFileNameChars()));
            string outputPath = Path.Combine(outputDir, $"{safeName}.png");

            // Render the first (and only) page of the sheet to a PNG file
            sheetRender.ToImage(0, outputPath);

            // Release resources used by SheetRender
            sheetRender.Dispose();
        }

        Console.WriteLine("All worksheets have been saved as separate PNG files.");
    }
}