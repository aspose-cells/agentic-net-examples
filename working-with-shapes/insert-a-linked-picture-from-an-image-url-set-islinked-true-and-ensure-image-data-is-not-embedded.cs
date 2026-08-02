// Title: Add a Linked Picture from a URL in Aspose.Cells for .NET (C#) – keep image unembedded
// Description: Creates a new workbook, inserts a picture from an external URL into cell B2 using Shapes.AddLinkedPicture, sets IsLink = true so the image is not embedded, and saves the file as LinkedPicture.xlsx.
// Keywords: Aspose.Cells AddLinkedPicture C# | linked picture URL Aspose.Cells | IsLink true Aspose.Cells | external image Excel .NET | prevent image embedding Aspose.Cells
// Common Searches: Aspose.Cells add linked picture from URL | How to set IsLink for pictures in Aspose.Cells | Create Excel file with external image C# | AddLinkedPicture without embedding data | Linked picture example Aspose.Cells .NET
// Developer Intent: Insert an external image as a linked picture in an Excel worksheet while ensuring the image data remains external.
// Use Cases: Generate reports that reference online logos to keep file size low. | Build templates that pull dynamic graphics from a web service at runtime. | Create workbooks where pictures update automatically when the source URL changes.
// AI Prompts: Show C# code to add a linked picture to a specific cell with IsLink set to true using Aspose.Cells. | Provide an example that inserts multiple linked pictures from a list of URLs without embedding image data. | Explain how to verify that a picture added with AddLinkedPicture is stored as a link in the workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsLinkedPictureExample
{
    // Creates a new workbook, inserts a picture from an external URL into cell B2 using Shapes.AddLinkedPicture, sets IsLink = true so the image is not embedded, and saves the file as LinkedPicture.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // URL of the image to be linked
                string imageUrl = "https://example.com/sample.jpg";

                // Add a linked picture to the worksheet (row 1, column 1, 100x100 pixels)
                // Rows and columns are zero‑based in Aspose.Cells, so (1,1) refers to cell B2
                Picture linkedPicture = worksheet.Shapes.AddLinkedPicture(1, 1, 100, 100, imageUrl);

                // Ensure the picture is marked as linked (IsLink = true) and no image data is embedded
                linkedPicture.IsLink = true;

                // Define output file path
                string outputPath = "LinkedPicture.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
