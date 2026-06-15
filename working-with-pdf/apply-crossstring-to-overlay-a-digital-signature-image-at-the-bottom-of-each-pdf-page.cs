using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace MyApp
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Path to the digital signature image that will be overlaid
                string signatureImagePath = "signature.png";

                // Verify that the signature image exists before using it
                if (!File.Exists(signatureImagePath))
                {
                    Console.WriteLine($"Signature image not found: {signatureImagePath}");
                }
                else
                {
                    // Iterate through each worksheet in the workbook
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Determine a row near the bottom of the sheet.
                        // MaxDataRow gives the last row that contains data; add a few rows for spacing.
                        int bottomRow = sheet.Cells.MaxDataRow + 2;

                        // Add the signature image at the calculated position (column 0 = first column)
                        int pictureIndex = sheet.Pictures.Add(bottomRow, 0, signatureImagePath);

                        // Retrieve the inserted picture object for optional adjustments
                        Picture signaturePicture = sheet.Pictures[pictureIndex];

                        // Example: keep the original size; you could also set Width/Height if needed
                        // signaturePicture.Width = 150;
                        // signaturePicture.Height = 50;
                    }
                }

                // Save the workbook as PDF (lifecycle rule: save)
                workbook.Save("SignedDocument.pdf", SaveFormat.Pdf);
                Console.WriteLine("Workbook saved as SignedDocument.pdf");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}