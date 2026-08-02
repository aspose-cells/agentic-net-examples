// Title: Unlock a protected Signature shape in Aspose.Cells (C#) to enable moving and resizing after confirmation
// Description: Demonstrates how to create a workbook, add a rectangle shape named "Signature", protect the worksheet, and then, when a confirmation flag is true, unlock the shape by clearing its IsLocked flag and the Move/Resize lock properties before saving the file.
// Keywords: Aspose.Cells | C# | unlock shape | shape lock | move resize | worksheet protection | Shape.IsLocked | SetLockedProperty | Signature shape | Excel protected sheet
// Common Searches: how to unlock a shape in a protected worksheet using Aspose.Cells | Aspose.Cells enable moving and resizing of a locked shape | C# unlock rectangle shape after worksheet protection | Aspose.Cells Shape.IsLocked false example | unlock signature shape in Excel with Aspose.Cells
// Developer Intent: Programmatically remove lock restrictions from a specific shape in a protected worksheet so it can be repositioned or resized after user approval.
// Use Cases: Allow a digital signature rectangle to be repositioned by the end‑user while the rest of the sheet stays protected. | Enable dynamic layout adjustments for a locked chart or image without unlocking the entire worksheet. | Keep a particular shape editable for form filling while maintaining overall sheet security.
// AI Prompts: Show C# code to unlock a shape's move and resize locks in Aspose.Cells after worksheet protection. | Provide an example that prompts the user and, upon confirmation, sets Shape.IsLocked and ShapeLockType to allow editing. | Explain how to protect an Excel sheet with Aspose.Cells but keep a specific shape editable.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSignatureUnlockDemo
{
    // Demonstrates how to create a workbook, add a rectangle shape named "Signature", protect the worksheet, and then, when a confirmation flag is true, unlock the shape by clearing its IsLocked flag and the Move/Resize lock properties before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a sample shape that will act as the "Signature"
            // Parameters: upper left row, upper left column, top, left, height, width
            Shape signatureShape = worksheet.Shapes.AddRectangle(2, 2, 50, 50, 150, 50);
            signatureShape.Name = "Signature";

            // Protect the worksheet so that locked objects cannot be moved/resized
            worksheet.Protect(ProtectionType.All);

            // Confirm unlocking (in real scenario this could be a user prompt)
            bool userConfirmed = true; // placeholder for confirmation logic

            if (userConfirmed)
            {
                // Unlock the shape itself
                signatureShape.IsLocked = false;

                // Additionally ensure that move and resize actions are not locked
                signatureShape.SetLockedProperty(ShapeLockType.Move, false);
                signatureShape.SetLockedProperty(ShapeLockType.Resize, false);
            }

            // Save the workbook (lifecycle: save)
            workbook.Save("SignatureUnlocked.xlsx");
        }
    }
}
