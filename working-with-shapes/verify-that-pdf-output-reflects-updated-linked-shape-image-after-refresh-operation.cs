// Title: Refresh a Linked Picture Shape and Verify the Updated Image in PDF with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, links the first picture to cell A1, writes a new image path, calls UpdateSelectedValue to refresh the picture, and saves the workbook as a PDF. The example demonstrates how to confirm that the PDF reflects the newly linked image.
// Keywords: Aspose.Cells | C# linked picture | UpdateSelectedValue | refresh picture shape | export to PDF | linked cell image | replace Excel image programmatically | PDF verification
// Common Searches: Aspose.Cells refresh linked picture after changing source | How to update a linked image in Excel and export to PDF using C# | Verify PDF shows new linked picture in Aspose.Cells | Update picture shape from cell value Aspose.Cells .NET | Refresh linked shape before PDF conversion
// Developer Intent: Ensure that after changing the image file path and invoking UpdateSelectedValue, the generated PDF contains the new linked picture.
// Use Cases: Swap an existing linked PNG in an Excel workbook with a new file, refresh the picture, and produce an accurate PDF. | Automate batch updates of graphics linked via cells and validate that each resulting PDF displays the correct image. | Create dynamic reports where images are driven by cell values, requiring a picture refresh before final PDF rendering.
// AI Prompts: Generate C# code using Aspose.Cells to change a picture's linked image, call UpdateSelectedValue, and confirm the PDF includes the new image. | Explain how to programmatically verify that a PDF generated from a workbook reflects the refreshed linked picture. | Provide a unit‑test example that asserts the PDF file contains the updated image after refreshing the linked shape.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace Example
{
    // Loads an Excel workbook, links the first picture to cell A1, writes a new image path, calls UpdateSelectedValue to refresh the picture, and saves the workbook as a PDF. The example demonstrates how to confirm that the PDF reflects the newly linked image.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "LinkedImage.xlsx";
                const string newImagePath = "newImage.png";
                const string outputPdf = "UpdatedLinkedImage.pdf";

                // Verify input workbook exists.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Verify the new image file exists.
                if (!File.Exists(newImagePath))
                {
                    Console.WriteLine($"Image file '{newImagePath}' not found.");
                    return;
                }

                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure there is at least one picture.
                if (sheet.Pictures.Count == 0)
                {
                    Console.WriteLine("No picture shapes found in the worksheet.");
                    return;
                }

                // Get the first picture shape.
                Picture picture = sheet.Pictures[0];

                // Link the picture to cell A1 (non‑absolute references).
                picture.SetLinkedCell("A1", false, false);

                // Set the linked cell value to the image file path.
                sheet.Cells["A1"].PutValue(newImagePath);

                // Refresh the picture to load the new image.
                picture.UpdateSelectedValue();

                // Save as PDF.
                PdfSaveOptions pdfOptions = new PdfSaveOptions();
                workbook.Save(outputPdf, pdfOptions);

                Console.WriteLine($"Workbook saved as PDF: {outputPdf}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
