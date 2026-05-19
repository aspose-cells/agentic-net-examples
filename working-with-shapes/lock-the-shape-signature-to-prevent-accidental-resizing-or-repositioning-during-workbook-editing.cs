using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class LockSignatureShape
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape that will act as the signature placeholder
        // Parameters: upper left row, upper left column, top offset, left offset, height, width
        Shape signature = worksheet.Shapes.AddRectangle(5, 2, 5, 2, 150, 80);
        signature.Text = "Signature";

        // Lock the shape so it cannot be modified when the sheet is protected
        signature.IsLocked = true;

        // Additionally lock specific actions: moving and resizing
        signature.SetLockedProperty(ShapeLockType.Move, true);
        signature.SetLockedProperty(ShapeLockType.Resize, true);

        // Protect the worksheet to enforce the lock
        worksheet.Protect(ProtectionType.All);

        // Save the workbook
        workbook.Save("LockedSignature.xlsx");
    }
}