// Title: Insert a JPEG image into a specific Excel cell and save the workbook as XLSX using Aspose.Cells for .NET
// AI Prompts: Generate C# code that checks for the existence of a JPEG file, adds it to cell B2 of a new workbook with Aspose.Cells, and saves the file as Result.xlsx. | Adapt the example to place the picture at cell D5, set its width and height to 200 px, and mark it as a background image before saving.
// Common Searches: how to add a JPEG picture to a particular cell in Excel with Aspose.Cells C# | Aspose.Cells C# insert image into worksheet and save as XLSX | C# verify image file path before inserting into Excel using Aspose.Cells | set picture size and position in Aspose.Cells workbook programmatically
// Tags: add picture to worksheet cell Aspose.Cells | insert JPEG image into Excel C# | Aspose.Cells picture positioning example | verify image file existence C# Aspose.Cells | save workbook as XLSX Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Loads a JPEG, validates its path, creates a new workbook, inserts the image into cell B2, and saves the workbook as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            string imagePath = @"C:\Images\sample.jpg";
            string outputPath = @"C:\Output\Result.xlsx";

            // Verify the image file exists
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add picture at cell B2 (row index 1, column index 1)
            int pictureIndex = sheet.Pictures.Add(1, 1, imagePath);

            // Retrieve the picture object (no need to set IsBackground; default is false)
            Picture picture = sheet.Pictures[pictureIndex];

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
