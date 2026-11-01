// Title: Export each worksheet of a CSV file to separate PNG images with Aspose.Cells for .NET (C#)
// Common Searches: c# how to export each worksheet of a csv file to individual png files using aspose.cells | aspose.cells convert csv workbook to png images per sheet c# example | render csv worksheets as png images with aspose.cells .net | save each sheet from a csv as a png image using aspose.cells c# code | asp.net aspose.cells export multiple worksheets to separate png files
// Tags: Aspose.Cells CSV to PNG conversion | C# SheetRender PNG export | ImageOrPrintOptions OnePagePerSheet usage | export multiple worksheets as images Aspose.Cells | Aspose.Cells render worksheet to PNG file

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads a CSV file into an Aspose.Cells Workbook, then iterates through all worksheets. For each sheet it configures ImageOrPrintOptions with SaveFormat.Png and OnePagePerSheet=true, creates a SheetRender, and saves the first page as a PNG file named Sheet1.png, Sheet2.png, etc. The code includes file‑existence checking and exception handling.
class Program
{
    static void Main()
    {
        // Path to the source CSV file
        string csvPath = "input.csv";

        // Verify that the CSV file exists to avoid FileNotFoundException
        if (!File.Exists(csvPath))
        {
            Console.WriteLine($"Error: The file \"{csvPath}\" was not found.");
            return;
        }

        try
        {
            // Load the CSV file into a workbook.
            // Aspose.Cells automatically creates a single worksheet for the CSV data.
            Workbook workbook = new Workbook(csvPath);

            // Iterate through all worksheets in the workbook.
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];

                // Configure image export options: PNG format, one page per sheet.
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    SaveFormat = SaveFormat.Png, // Correct property for image format
                    OnePagePerSheet = true
                };

                // Render the worksheet to an image.
                SheetRender sheetRender = new SheetRender(sheet, imgOptions);

                // Export the first (and only) page of the sheet to a PNG file.
                // File name includes the worksheet index (1‑based) for uniqueness.
                string pngPath = $"Sheet{i + 1}.png";
                sheetRender.ToImage(0, pngPath);
            }

            Console.WriteLine("Export completed.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
