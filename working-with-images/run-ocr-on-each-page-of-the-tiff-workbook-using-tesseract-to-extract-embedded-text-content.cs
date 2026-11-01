// Title: How to embed the first page of a multi‑page TIFF as a free‑floating picture in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells insert first page of multi‑page TIFF into Excel worksheet | Add TIFF image as free‑floating picture in Excel with Aspose.Cells .NET | Save Excel workbook with embedded TIFF picture using Aspose.Cells C# | Aspose.Cells picture placement freefloating example in C# | Load multi‑page TIFF and insert into workbook using Aspose.Cells
// Tags: Aspose.Cells insert TIFF picture | C# freefloating picture placement | multi-page TIFF to Excel worksheet | Aspose.Cells save workbook with image | picture.Add method Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample checks for the TIFF file, creates the output folder, instantiates a Workbook, selects the first worksheet, inserts the first page of the multi‑page TIFF as a picture using Pictures.Add, sets the picture's Placement to FreeFloating, and saves the workbook as an .xlsx file.
class TiffImageProcessor
{
    // Path to the multi‑page TIFF file
    private const string TiffFilePath = @"C:\Path\To\Your\Document.tif";

    // Path where the resulting Excel workbook will be saved
    private const string OutputWorkbookPath = @"C:\Path\To\Your\Document.xlsx";

    static void Main()
    {
        try
        {
            // Verify that the TIFF file exists
            if (!File.Exists(TiffFilePath))
            {
                Console.WriteLine($"TIFF file not found: {TiffFilePath}");
                return;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(OutputWorkbookPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Use the default first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Page_1";

            // Insert the TIFF file as a picture (first page) into the worksheet
            // Using the overload that accepts a file path avoids System.Drawing dependencies
            int pictureIndex = sheet.Pictures.Add(0, 0, TiffFilePath);
            Picture picture = sheet.Pictures[pictureIndex];
            picture.Placement = PlacementType.FreeFloating; // Allow free positioning

            // Save the workbook to the specified path
            workbook.Save(OutputWorkbookPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to: {OutputWorkbookPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
