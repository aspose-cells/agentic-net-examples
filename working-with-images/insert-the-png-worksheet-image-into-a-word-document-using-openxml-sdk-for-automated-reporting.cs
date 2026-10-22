// Title: Insert a worksheet PNG image into a Word document with OpenXML SDK and Aspose.Cells in C#
// AI Prompts: Write C# code that loads an .xlsx file using Aspose.Cells, renders the first worksheet to a PNG in a memory stream, creates a .docx with OpenXML SDK, and embeds the PNG into the Word document body. | Generate a method that takes Excel and Word file paths, uses SheetRender to produce a PNG stream, adds the image as an ImagePart to a WordprocessingDocument, and saves the resulting report.
// Common Searches: C# how to embed a PNG generated from an Excel worksheet into a Word document using OpenXML | Aspose.Cells export first sheet as image and insert into .docx programmatically | automate report creation by adding Excel sheet screenshot to Word with OpenXML SDK | insert worksheet image into Word using C# and OpenXML without saving temporary file | render Excel worksheet to memory stream and embed in Word document in .NET
// Tags: Aspose.Cells render worksheet to PNG | OpenXML SDK embed image in docx | C# export Excel sheet as image | automated reporting using Excel image in Word | SheetRender memory stream to Word image part

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example demonstrates loading an Excel workbook, rendering the first worksheet to a PNG using Aspose.Cells, then creating a WordprocessingDocument with OpenXML SDK and inserting the PNG as an image part into the Word document, enabling fully automated report generation.
class WorksheetImageExport
{
    static void Main()
    {
        try
        {
            // Paths for source Excel and output image
            string excelPath = @"C:\Temp\SourceWorkbook.xlsx";
            string imagePath = @"C:\Temp\WorksheetImage.png";

            // Verify that the source Excel file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: Excel file not found at '{excelPath}'.");
                return;
            }

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(imagePath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(excelPath);
            Worksheet sheet = workbook.Worksheets[0]; // first worksheet

            // Set image export options – default format is PNG, one page per sheet
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true,
                Transparent = false
            };

            // Render the worksheet to an image stored in a memory stream
            using (MemoryStream imageStream = new MemoryStream())
            {
                SheetRender renderer = new SheetRender(sheet, imgOptions);
                renderer.ToImage(0, imageStream); // render first page
                imageStream.Position = 0; // reset stream position

                // Save the image to a file
                using (FileStream fileStream = new FileStream(imagePath, FileMode.Create, FileAccess.Write))
                {
                    imageStream.CopyTo(fileStream);
                }
            }

            Console.WriteLine($"Worksheet image saved to '{imagePath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
