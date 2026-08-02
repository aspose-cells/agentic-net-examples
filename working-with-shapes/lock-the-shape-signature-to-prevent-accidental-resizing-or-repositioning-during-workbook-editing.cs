// Title: Lock a shape to prevent moving or resizing in Aspose.Cells for .NET
// Description: Shows how to create a workbook, add a rectangle shape as a signature placeholder, set IsLocked and specific ShapeLockType flags for movement and resizing, protect the worksheet, and save the file so the shape stays fixed during editing.
// Keywords: Aspose.Cells lock shape | C# lock Excel shape | ShapeLockType Move | ShapeLockType Resize | IsLocked property Aspose.Cells | protect worksheet Aspose.Cells | signature placeholder Excel | prevent shape editing Aspose.Cells | global
// Common Searches: How to lock a shape in Aspose.Cells C# | Aspose.Cells prevent shape resizing | Lock signature shape in Excel using Aspose.Cells | Set shape lock properties .NET | Worksheet protection lock shapes Aspose
// Developer Intent: Prevent a shape from being moved or resized after the worksheet is protected in an Aspose.Cells workbook.
// Use Cases: Create a fixed‑position signature placeholder that cannot be altered by end users. | Secure a company logo or watermark shape so its size and location remain unchanged in a protected sheet. | Lock form‑field shapes in a data‑entry worksheet to maintain layout integrity.
// AI Prompts: Generate C# code that locks a shape's movement and resizing in Aspose.Cells and protects the worksheet. | Explain the difference between IsLocked and ShapeLockType settings for shapes in Aspose.Cells. | Provide an example of locking multiple shapes with different lock types using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, add a rectangle shape as a signature placeholder, set IsLocked and specific ShapeLockType flags for movement and resizing, protect the worksheet, and save the file so the shape stays fixed during editing.
class LockSignatureShapeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape that will act as the signature placeholder
        // Parameters: upper left row, upper left column, top offset, left offset, height, width
        Shape signatureShape = worksheet.Shapes.AddRectangle(2, 1, 0, 0, 100, 200);
        signatureShape.Name = "Signature";
        signatureShape.Text = "Signature";

        // Lock the shape so it cannot be moved or resized when the sheet is protected
        signatureShape.IsLocked = true; // General lock
        signatureShape.SetLockedProperty(ShapeLockType.Move, true);    // Prevent moving
        signatureShape.SetLockedProperty(ShapeLockType.Resize, true);  // Prevent resizing

        // Protect the worksheet (all protection types) to enforce the lock
        worksheet.Protect(ProtectionType.All);

        // Save the workbook
        workbook.Save("LockedSignatureShape.xlsx");
    }
}
