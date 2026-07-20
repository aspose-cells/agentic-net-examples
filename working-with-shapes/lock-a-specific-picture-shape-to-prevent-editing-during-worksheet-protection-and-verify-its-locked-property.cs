// Title: Lock Picture Shape in Aspose.Cells for .NET, Protect Worksheet & Verify Locked State
// Description: C# example that creates a workbook, adds a PNG picture, sets the picture's IsLocked flag and a specific ShapeLockType (Move), protects the worksheet, prints the lock status, and saves the file as LockedPictureDemo.xlsx. Includes fallback code to generate a placeholder image if the source file is missing.
// Keywords: Aspose.Cells lock picture | C# picture shape protection | IsLocked property Aspose.Cells | ShapeLockType Move | worksheet protection objects | verify picture lock status | .NET Excel image lock
// Common Searches: how to lock a picture in Aspose.Cells .NET | prevent editing of images after worksheet protection | check IsLocked flag for picture shape | set ShapeLockType for Excel picture using Aspose.Cells | Aspose.Cells protect sheet objects example
// Developer Intent: Secure a picture shape from editing (move, resize, delete) by locking it, apply worksheet protection, and programmatically confirm the lock settings.
// Use Cases: Insert a company logo into a generated report and lock it so end‑users cannot alter or remove it. | Create a template with fixed graphics, protect the sheet, and validate that the graphics remain immutable. | Automate verification of shape lock settings after applying worksheet protection in a CI pipeline.
// AI Prompts: Write C# code with Aspose.Cells to add a PNG, lock its movement, protect the worksheet, and output the lock status. | Explain the relationship between IsLocked, SetLockedProperty, and worksheet protection for picture shapes in Aspose.Cells. | Provide a step‑by‑step guide to generate a placeholder PNG when the image file is absent before inserting it into a workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that creates a workbook, adds a PNG picture, sets the picture's IsLocked flag and a specific ShapeLockType (Move), protects the worksheet, prints the lock status, and saves the file as LockedPictureDemo.xlsx. Includes fallback code to generate a placeholder image if the source file is missing.
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
            string picturePath = "sample.png";

            // Ensure the picture file exists; create a minimal PNG if it does not
            if (!File.Exists(picturePath))
            {
                try
                {
                    // 1x1 pixel transparent PNG (base64 encoded)
                    const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9WcAAAAASUVORK5CYII=";
                    byte[] pngBytes = Convert.FromBase64String(base64Png);
                    File.WriteAllBytes(picturePath, pngBytes);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Failed to create placeholder image: " + ex.Message);
                    return;
                }
            }

            // Add the picture to the worksheet
            int pictureIndex = worksheet.Pictures.Add(1, 1, picturePath);
            Picture picture = worksheet.Pictures[pictureIndex];

            // Lock the picture shape so it cannot be edited when the sheet is protected
            picture.IsLocked = true;

            // Optionally lock specific actions, e.g., moving the picture
            picture.SetLockedProperty(ShapeLockType.Move, true);

            // Protect the worksheet (default settings disallow editing objects)
            worksheet.Protect(ProtectionType.All);

            // Verify that the picture is locked
            Console.WriteLine("Picture.IsLocked: " + picture.IsLocked);
            Console.WriteLine("Picture Move locked: " + picture.GetLockedProperty(ShapeLockType.Move));

            // Save the workbook
            workbook.Save("LockedPictureDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
