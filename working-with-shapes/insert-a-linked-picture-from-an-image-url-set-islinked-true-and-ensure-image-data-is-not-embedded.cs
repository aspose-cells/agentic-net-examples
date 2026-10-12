// Title: Insert a linked picture from an external URL into an Excel worksheet using Aspose.Cells for .NET without embedding the image
// AI Prompts: Add a picture to cell A1 that references https://example.com/image.png and set its IsLinked property to true using Aspose.Cells. | Save the workbook as an XLSX file while preserving the picture as a linked image rather than embedding the binary data. | Verify that the generated Excel file contains a linked picture by inspecting the workbook's picture collection.
// Common Searches: Aspose.Cells C# add picture from URL as linked image | How to keep Excel picture linked instead of embedded using Aspose.Cells | Save workbook with external image reference in .NET Aspose.Cells | Set IsLinked property for picture in Aspose.Cells example | Create Excel file with web-linked picture using Aspose.Cells for .NET
// Tags: add linked picture Aspose.Cells | picture IsLinked property .NET | save workbook with external image Aspose.Cells | Excel linked image from URL C# | prevent image embedding Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a new workbook, inserts a picture at cell A1 that points to an external URL, marks it as linked (IsLinked = true), and saves the file as an XLSX workbook, ensuring the image data remains external and is not embedded.
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

            // URL of the image to link
            string imageUrl = "https://example.com/image.png";

            // Add a picture at cell A1 (row 0, column 0). The picture will be linked if the source is a URL.
            int pictureIndex = sheet.Pictures.Add(0, 0, imageUrl);

            // Retrieve the picture object (optional, for further manipulation)
            Picture picture = sheet.Pictures[pictureIndex];

            // Define output file path
            string outputPath = "LinkedPicture.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
