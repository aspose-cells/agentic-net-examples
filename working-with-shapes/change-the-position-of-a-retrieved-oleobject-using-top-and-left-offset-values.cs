using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ChangeOleObjectPosition
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add an OLE object to the worksheet (using an image as the icon)
        // Ensure "sample.png" exists in the executable directory
        byte[] imageData = File.ReadAllBytes("sample.png");
        int oleIndex = worksheet.OleObjects.Add(5, 5, 100, 100, imageData);

        // Retrieve the added OLE object
        OleObject oleObject = worksheet.OleObjects[oleIndex];

        // Define the offset values (in pixels)
        int topOffset = 20;   // move down by 20 pixels
        int leftOffset = 30;  // move right by 30 pixels

        // Adjust the position using the Top and Left properties
        oleObject.Top += topOffset;
        oleObject.Left += leftOffset;

        // Save the workbook with the updated OLE object position
        workbook.Save("OleObjectMoved.xlsx");
    }
}