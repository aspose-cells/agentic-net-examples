using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class UnlockSignatureShape
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape that will act as the signature
        // Parameters: upper left row, upper left column, offsetX, offsetY, width, height
        Shape signature = sheet.Shapes.AddRectangle(5, 2, 0, 0, 150, 50);
        signature.Name = "Signature";

        // Protect the worksheet to simulate a locked environment
        sheet.Protect(ProtectionType.All);

        // Display the initial lock state
        Console.WriteLine("Initial IsLocked: " + signature.IsLocked);

        // Simulate user confirmation (replace with real confirmation logic as needed)
        bool userConfirmed = true;

        if (userConfirmed)
        {
            // Unlock the shape itself
            signature.IsLocked = false;

            // Unlock specific properties that affect size and position
            signature.SetLockedProperty(ShapeLockType.Move, false);
            signature.SetLockedProperty(ShapeLockType.Resize, false);
        }

        // Verify the lock states after unlocking
        Console.WriteLine("After unlocking IsLocked: " + signature.IsLocked);
        Console.WriteLine("Move locked: " + signature.GetLockedProperty(ShapeLockType.Move));
        Console.WriteLine("Resize locked: " + signature.GetLockedProperty(ShapeLockType.Resize));

        // Save the workbook
        workbook.Save("UnlockedSignature.xlsx");
    }
}