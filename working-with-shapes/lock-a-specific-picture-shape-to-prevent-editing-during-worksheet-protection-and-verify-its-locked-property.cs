// Title: How to lock a picture shape in an Aspose.Cells worksheet and verify its IsLocked property using C#
// AI Prompts: Insert a picture into a worksheet, set its IsLocked flag to true, protect the sheet with all options, and print the lock status. | Update the sample to gracefully handle a missing image file and confirm that the picture remains locked after worksheet protection.
// Common Searches: C# Aspose.Cells lock inserted image from editing when worksheet is protected | Aspose.Cells picture.IsLocked verification after protecting a sheet | prevent users from moving or resizing pictures in a protected Excel worksheet using Aspose.Cells .NET | how to check if a picture shape is locked in an Aspose.Cells workbook | example code for locking a picture shape and protecting a worksheet in Aspose.Cells
// Tags: lock picture shape Aspose.Cells | picture.IsLocked property C# | protect worksheet Aspose.Cells | add picture to worksheet Aspose.Cells | verify picture lock status Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a new workbook, adds a picture from a file to cell B2, sets the picture's IsLocked property to true, prints the lock state, protects the worksheet with all protection options, and saves the file as LockedPicture.xlsx while handling missing image files.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the image file
            string imagePath = "sample.png";

            // Ensure the image file exists before adding it to the worksheet
            if (File.Exists(imagePath))
            {
                // Add a picture shape to the worksheet (row 1, column 1 – zero‑based index)
                int pictureIndex = worksheet.Pictures.Add(1, 1, imagePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Lock the picture shape so it cannot be edited when the sheet is protected
                picture.IsLocked = true;

                // Verify the Locked property of the picture
                Console.WriteLine("Picture Locked: " + picture.IsLocked);
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}. Skipping picture insertion.");
            }

            // Protect the worksheet (all protection options enabled)
            worksheet.Protect(ProtectionType.All);

            // Save the workbook
            string outputPath = "LockedPicture.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
