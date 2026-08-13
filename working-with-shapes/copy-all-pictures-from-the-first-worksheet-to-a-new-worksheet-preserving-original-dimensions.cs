// Title: Copy all pictures from the first worksheet to a new sheet preserving size with Aspose.Cells for .NET
// Description: Loads a source workbook, creates a new workbook, and copies every picture from the first worksheet to a newly added worksheet. The code records each picture's cell coordinates, inserts a 1×1 transparent PNG placeholder, then uses Picture.Copy with default CopyOptions to retain the original dimensions and position before saving the result.
// Keywords: Aspose.Cells copy pictures | C# copy images between worksheets | preserve picture dimensions Excel | Aspose.Cells placeholder PNG | transfer shapes .NET workbook | copy pictures Aspose.Cells example
// Common Searches: how to copy pictures between worksheets using Aspose.Cells | preserve image size when moving Excel pictures programmatically | Aspose.Cells copy all images from first sheet to new sheet | C# copy Excel pictures with original dimensions | Aspose.Cells picture.Copy example
// Developer Intent: Programmatically duplicate every picture from the first worksheet into a new worksheet while keeping the original size and placement.
// Use Cases: Create a visual replica of a template sheet for archiving without altering layout. | Generate language‑specific reports by reusing branding images from a master workbook. | Build a summary workbook that mirrors the graphics of a source sheet for presentation purposes.
// AI Prompts: Write C# code with Aspose.Cells that copies all pictures from the first worksheet to a new worksheet, preserving their dimensions and positions. | Explain why a transparent PNG placeholder stream is required when using Picture.Copy in Aspose.Cells. | Show an alternative approach to copy pictures between worksheets without using a placeholder image.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPictureCopyDemo
{
    // Loads a source workbook, creates a new workbook, and copies every picture from the first worksheet to a newly added worksheet. The code records each picture's cell coordinates, inserts a 1×1 transparent PNG placeholder, then uses Picture.Copy with default CopyOptions to retain the original dimensions and position before saving the result.
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "source.xlsx";
                const string destPath = "copied_pictures.xlsx";

                // Verify source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file \"{sourcePath}\" not found.");
                    return;
                }

                // Load the source workbook (assumed to contain pictures in the first worksheet)
                Workbook sourceWorkbook = new Workbook(sourcePath);
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

                // Create a new workbook and get its first worksheet
                Workbook destWorkbook = new Workbook();
                Worksheet destSheet = destWorkbook.Worksheets[0];
                destSheet.Name = "CopiedPictures";

                // Prepare copy options (default options are sufficient for pictures)
                CopyOptions copyOptions = new CopyOptions();

                // Iterate through all pictures in the source worksheet
                for (int i = 0; i < sourceSheet.Pictures.Count; i++)
                {
                    Picture srcPic = sourceSheet.Pictures[i];

                    // Preserve the original picture position and size by using the same cell coordinates
                    int topRow = srcPic.UpperLeftRow;
                    int leftColumn = srcPic.UpperLeftColumn;
                    int bottomRow = srcPic.LowerRightRow;
                    int rightColumn = srcPic.LowerRightColumn;

                    // Create a 1x1 transparent PNG in memory to serve as a placeholder
                    byte[] transparentPng = new byte[]
                    {
                        0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                        0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                        0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
                        0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
                        0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
                        0x54,0x78,0x9C,0x63,0x00,0x01,0x00,0x00,
                        0x05,0x00,0x01,0x0D,0x0A,0x2D,0xB4,0x00,
                        0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
                        0x42,0x60,0x82
                    };

                    using (MemoryStream ms = new MemoryStream(transparentPng))
                    {
                        // Add placeholder picture to destination sheet
                        int placeholderIndex = destSheet.Pictures.Add(topRow, leftColumn, bottomRow, rightColumn, ms);
                        Picture destPic = destSheet.Pictures[placeholderIndex];

                        // Copy the source picture into the placeholder picture, preserving dimensions
                        destPic.Copy(srcPic, copyOptions);
                    }
                }

                // Save the destination workbook with the copied pictures
                destWorkbook.Save(destPath);
                Console.WriteLine($"Pictures copied successfully to \"{destPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
