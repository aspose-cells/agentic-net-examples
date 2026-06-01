using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

class SetOleObjectVisibility
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add an OLE object with default icon (null image data)
            int oleIndex = sheet.OleObjects.Add(5, 2, 100, 100, (byte[])null);
            OleObject ole = sheet.OleObjects[oleIndex];

            // Hide the OLE object so it will not be visible in the worksheet
            ole.IsHidden = true;

            // Save the workbook
            workbook.Save("OleObjectHidden.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}