// Title: Lock a Picture Shape and Verify Its Locked Property in Aspose.Cells for .NET
// Description: Demonstrates how to add a picture to a worksheet, set its IsLocked flag, apply a move lock via ShapeLockType.Move, protect the sheet with all protection types, output the lock status, and save the workbook using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | lock picture shape | IsLocked property | ShapeLockType.Move | protect worksheet | Excel picture lock | shape protection | read‑only image | Aspose.Cells example
// Common Searches: Aspose.Cells lock picture shape C# | How to prevent editing of an image in Excel with Aspose.Cells | Check IsLocked status of a picture in Aspose.Cells | Protect worksheet and lock shapes Aspose.Cells .NET | Set ShapeLockType.Move for picture Aspose.Cells
// Developer Intent: The developer needs to make a picture shape immutable after worksheet protection and confirm that the lock is active.
// Use Cases: Insert a corporate logo, lock its position and size, and protect the template so users cannot modify it. | Apply move, size, and delete locks to multiple images before distributing a read‑only workbook. | Validate that previously locked picture shapes remain protected after re‑applying worksheet protection.
// AI Prompts: Write C# code with Aspose.Cells that adds a picture, locks its movement, protects the sheet, and prints the lock state. | Show how to lock and verify move, size, and delete properties of a picture shape using Aspose.Cells for .NET. | Explain the steps to unlock a picture shape after a worksheet has been protected with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to add a picture to a worksheet, set its IsLocked flag, apply a move lock via ShapeLockType.Move, protect the sheet with all protection types, output the lock status, and save the workbook using Aspose.Cells for C#.
    public class LockPictureShapeDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the picture file
            string picturePath = "example.jpg";

            // Add a picture to the worksheet if the file exists
            if (File.Exists(picturePath))
            {
                // Parameters: upper left row, upper left column, picture file name
                int pictureIndex = worksheet.Pictures.Add(2, 1, picturePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Lock the picture shape so it cannot be edited when the sheet is protected
                picture.IsLocked = true; // Shape.IsLocked property

                // Optionally lock specific properties (e.g., moving the picture)
                picture.SetLockedProperty(ShapeLockType.Move, true);

                // Protect the worksheet (all protection types)
                worksheet.Protect(ProtectionType.All);

                // Verify the locked state
                Console.WriteLine("Picture IsLocked property: " + picture.IsLocked);
                bool isMoveLocked = picture.GetLockedProperty(ShapeLockType.Move);
                Console.WriteLine("Picture Move locked: " + isMoveLocked);
            }
            else
            {
                Console.WriteLine($"Picture file '{picturePath}' not found. Skipping picture insertion.");
                // Still protect the worksheet even without a picture
                worksheet.Protect(ProtectionType.All);
            }

            // Save the workbook (save rule)
            string outputPath = "LockedPictureDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
