using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class LockPictureShapeDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the picture file
            string picturePath = "example.jpg";

            // Add picture only if the file exists
            if (File.Exists(picturePath))
            {
                int pictureIndex = worksheet.Pictures.Add(1, 1, picturePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Lock the picture so it cannot be edited when the sheet is protected
                picture.IsLocked = true;

                // Protect the worksheet (all protection options)
                worksheet.Protect(ProtectionType.All);

                // Verify that the picture is locked
                Console.WriteLine("Picture IsLocked property: " + picture.IsLocked);
                Console.WriteLine("Shape GetLockedProperty for Move: " + picture.GetLockedProperty(ShapeLockType.Move));
            }
            else
            {
                Console.WriteLine($"Picture file not found: {picturePath}. Skipping picture insertion.");
                // Still protect the worksheet even without a picture
                worksheet.Protect(ProtectionType.All);
            }

            // Save the workbook
            workbook.Save("LockedPictureDemo.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}