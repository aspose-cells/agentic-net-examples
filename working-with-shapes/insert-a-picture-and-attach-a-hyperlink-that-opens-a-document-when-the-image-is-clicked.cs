// Title: Insert an image into an Excel worksheet and attach a hyperlink that opens a PDF using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# insert picture and link it to a PDF file | Add clickable image to Excel workbook using Aspose.Cells .NET | Set hyperlink on a picture in an Excel sheet with Aspose.Cells API | Create tooltip for picture hyperlink in Aspose.Cells workbook | Open external document from an Excel image using Aspose.Cells C#
// Tags: insert picture with hyperlink Aspose.Cells | picture hyperlink to external PDF Aspose.Cells | add screen tip to Excel image Aspose.Cells | alternative text for picture Aspose.Cells | save workbook containing linked image Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace InsertPictureWithHyperlink
{
    // The example creates a new workbook, inserts a PNG image at cell B2, assigns a hyperlink that opens a PDF file, sets a screen tip and alternative text for the picture, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Paths – adjust these to point to existing files on your machine.
                string imagePath = "sampleImage.png";               // The picture to insert.
                string targetDocumentPath = "TargetDocument.pdf";   // Document to open when picture is clicked.
                string outputPath = "WorkbookWithPicture.xlsx";

                // Validate input files.
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                    return;
                }

                if (!File.Exists(targetDocumentPath))
                {
                    Console.WriteLine($"Target document not found: {targetDocumentPath}");
                    return;
                }

                // Create a new workbook.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Demo";

                // Insert the picture at cell B2 (row index 1, column index 1).
                int pictureIndex = sheet.Pictures.Add(1, 1, imagePath);
                Picture picture = sheet.Pictures[pictureIndex];

                // Attach a hyperlink that points to the target document.
                picture.Hyperlink.Address = targetDocumentPath;
                picture.Hyperlink.ScreenTip = "Click to open the document";

                // Optional: set the tooltip that appears when hovering over the picture.
                picture.AlternativeText = "Click to open the document";

                // Save the workbook.
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved to '{outputPath}'.");
                Console.WriteLine("The picture now contains a hyperlink that opens the document when clicked.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
