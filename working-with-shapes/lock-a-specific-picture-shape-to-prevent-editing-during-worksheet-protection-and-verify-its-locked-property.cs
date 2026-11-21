// Title: Lock a picture shape in an Excel worksheet using Aspose.Cells for .NET and verify its IsLocked property
// Common Searches: Aspose.Cells C# lock picture shape on protected worksheet | prevent editing of inserted image in Excel using Aspose.Cells | check IsLocked property of picture after worksheet protection in .NET | add PNG as picture and lock it with Aspose.Cells workbook | protect Excel sheet while keeping picture locked using Aspose.Cells
// Tags: Aspose.Cells picture shape locking | worksheet protection image lock | C# Aspose.Cells add picture | IsLocked property verification | Excel workbook save locked picture

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads a PNG file, inserts it as a picture into the first worksheet of a new workbook, sets the picture's IsLocked property to true, protects the worksheet with all protection types, prints the lock status, and saves the workbook as LockedPicture.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Determine the path of the image file relative to the executable directory
            string exeDir = AppDomain.CurrentDomain.BaseDirectory;
            string imagePath = Path.Combine(exeDir, "sample.png");

            // Verify that the image file exists to avoid FileNotFoundException
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a picture shape to the worksheet
            int pictureIndex = worksheet.Pictures.Add(1, 1, imagePath);
            Picture picture = worksheet.Pictures[pictureIndex];

            // Lock the picture shape so it cannot be edited when the sheet is protected
            picture.IsLocked = true;

            // Protect the worksheet (all protection types enabled)
            worksheet.Protect(ProtectionType.All);

            // Verify that the picture's lock status is set to true
            Console.WriteLine("Picture IsLocked property: " + picture.IsLocked);

            // Save the workbook
            string outputPath = Path.Combine(exeDir, "LockedPicture.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
