using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Put a value into a cell that will be linked to the picture
        sheet.Cells["B2"].PutValue("Linked Text");

        // Add a picture to the worksheet (adjust the file path as needed)
        int pictureIdx = sheet.Pictures.Add(2, 2, "sample.png"); // row, column, image file
        Picture picture = sheet.Pictures[pictureIdx];

        // Link the picture to the cell B2
        // Parameters: linkedCell, isRowAbsolute, isColumnAbsolute
        picture.SetLinkedCell("B2", false, false);

        // Refresh the picture so it reflects the current value of the linked cell
        picture.UpdateSelectedValue();

        // Save the workbook
        workbook.Save("UpdatedPicture.xlsx");
    }
}