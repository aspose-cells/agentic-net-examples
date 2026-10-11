// Title: Lock a picture shape in an Excel worksheet with Aspose.Cells for .NET and verify its IsLocked property
// AI Prompts: Add a PNG file to a worksheet, set its IsLocked property to true, protect the sheet with all options, and save the workbook. | Read and print the IsLocked flag of a picture after the worksheet has been protected. | Insert several picture shapes, lock each one, apply worksheet protection, and output the lock status for every picture.
// Common Searches: asp.net how to lock an inserted picture in Excel using Aspose.Cells | c# verify picture IsLocked after worksheet protection Aspose.Cells example | prevent editing of images in a protected Excel workbook with Aspose.Cells for .NET | set picture.IsLocked before protecting sheet Aspose.Cells C# code | example of locking picture shapes in Excel file using Aspose.Cells library
// Tags: picture shape lock Aspose.Cells | worksheet protection image IsLocked | insert PNG protect sheet C# | verify picture lock status Aspose.Cells | prevent image editing Excel Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// // Creates a new workbook, inserts a PNG picture at cell B3, sets picture.IsLocked = true, protects the worksheet with all protection options, prints the lock status to the console, and saves the file as LockedPicture.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the image file to be inserted
            string imagePath = "sample.png";

            // Ensure the image file exists to avoid FileNotFoundException
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add the picture to the worksheet at row 2, column 2 (zero‑based indices)
            int pictureIndex = worksheet.Pictures.Add(2, 2, imagePath);
            Picture picture = worksheet.Pictures[pictureIndex];

            // Lock the picture so it cannot be edited when the sheet is protected
            picture.IsLocked = true;

            // Protect the worksheet with all protection options
            worksheet.Protect(ProtectionType.All);

            // Output verification
            Console.WriteLine("Picture IsLocked property: " + picture.IsLocked);

            // Save the workbook
            string outputPath = "LockedPicture.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
