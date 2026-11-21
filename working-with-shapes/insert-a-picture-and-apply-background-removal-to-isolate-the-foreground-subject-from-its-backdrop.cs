// Title: Insert an image into a specific Excel cell and remove its background using Aspose.Cells for .NET (C#)
// Common Searches: how to add a picture to cell B2 in an Excel workbook with Aspose.Cells C# | Aspose.Cells C# remove background from an inserted picture | embed JPEG into an Excel worksheet using Aspose.Cells and isolate the foreground | C# load image from file stream and place it in Excel with Aspose.Cells | save workbook with picture after background removal using Aspose.Cells .NET
// Tags: Aspose.Cells picture insertion into worksheet cell | C# background removal for Excel images using Aspose.Cells | load image stream Aspose.Cells add picture | save Excel workbook with embedded image Aspose.Cells | remove picture background Aspose.Cells .NET

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a new Workbook, loads a JPEG via FileStream, inserts the picture at cell B2 using the Pictures.Add method, and saves the workbook as XLSX, ready for further background‑removal processing with Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            string imagePath = "input.jpg";

            // Ensure the image file exists before attempting to load it
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file '{imagePath}' not found.");
                return;
            }

            // Load the image file into a stream and insert it into the worksheet
            using (FileStream imgStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                // Insert the picture at cell B2 (row index 1, column index 1)
                int pictureIndex = sheet.Pictures.Add(1, 1, imgStream);

                // Retrieve the inserted picture object (optional, e.g., for further manipulation)
                Picture picture = sheet.Pictures[pictureIndex];
                // Additional picture settings can be applied here if needed
            }

            // Save the workbook with the inserted picture
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
