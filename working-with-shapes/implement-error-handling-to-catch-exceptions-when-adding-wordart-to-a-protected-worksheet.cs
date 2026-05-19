using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AddWordArtProtectedWorksheet
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Disallow editing of drawing objects and protect the worksheet
        sheet.Protection.AllowEditingObject = false;
        sheet.Protect(ProtectionType.All);

        // Attempt to add WordArt to the protected worksheet
        try
        {
            ShapeCollection shapes = sheet.Shapes;
            Shape wordArt = shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle3, // preset style
                "Protected WordArt",               // text
                2,  // topRow (upper left row index)
                0,  // top offset in pixels
                2,  // leftColumn (upper left column index)
                0,  // left offset in pixels
                100, // height in pixels
                400  // width in pixels
            );

            Console.WriteLine("WordArt added successfully.");
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.Shape)
        {
            // Expected exception when trying to add a shape to a protected sheet
            Console.WriteLine("Error adding WordArt to a protected worksheet: " + ex.Message);
        }
        catch (Exception ex)
        {
            // Any other unexpected errors
            Console.WriteLine("Unexpected error: " + ex.Message);
        }

        // Save the workbook
        workbook.Save("ProtectedWordArtDemo.xlsx");
    }
}