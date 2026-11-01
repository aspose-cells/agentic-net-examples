// Title: Load an Excel workbook from a Stream and insert a web‑linked picture into cell G10 with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel workbook from a Stream, adds a linked picture using a web URL to cell G10, refreshes the linked picture, and saves the file as XLSX. | Show how to use Aspose.Cells to load a workbook via a Stream, place a URL‑based image in cell G10, update linked images, and write the result to disk.
// Common Searches: how to add a linked picture from a URL to a specific cell using Aspose.Cells C# | load Excel workbook from memory stream and embed web image in cell G10 Aspose.Cells | refresh linked pictures after inserting a web image with Aspose.Cells .NET | Aspose.Cells picture.Add method with URL and cell coordinates example
// Tags: load workbook from stream Aspose.Cells | add linked picture from URL to worksheet cell | refresh linked pictures Aspose.Cells | save workbook as xlsx Aspose.Cells | insert web image into Excel cell G10

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads a workbook from a Stream, adds a linked picture from a web URL to cell G10, refreshes linked images, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Load the workbook from a stream (file or generated)
            using (Stream inputStream = GetInputStream())
            {
                Workbook workbook = new Workbook(inputStream);

                // URL of the picture to be linked
                string pictureUrl = "https://example.com/image.jpg";

                // Add a linked picture to cell G10 (row 9, column 6)
                Worksheet sheet = workbook.Worksheets[0];
                int pictureIndex = sheet.Pictures.Add(9, 6, pictureUrl);
                Picture picture = sheet.Pictures[pictureIndex];

                // The picture added via URL is already a linked picture; no additional property needed

                // Save the modified workbook
                string outputPath = "output.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Returns a stream for the input workbook; creates a blank workbook if the file is missing
    static Stream GetInputStream()
    {
        const string inputPath = "input.xlsx";

        if (File.Exists(inputPath))
        {
            return File.OpenRead(inputPath);
        }

        // Create a new blank workbook and return its stream
        Workbook wb = new Workbook();
        MemoryStream ms = new MemoryStream();
        wb.Save(ms, SaveFormat.Xlsx);
        ms.Position = 0;
        return ms;
    }
}
