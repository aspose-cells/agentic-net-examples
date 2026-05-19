using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class SetPictureAnchor
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells B2:D4 (rows 1-3, columns 1-3)
        worksheet.Cells.Merge(1, 1, 3, 3);

        // Add a picture to the worksheet (adjust the path to an existing image)
        int pictureIndex = worksheet.Pictures.Add(1, 1, "sample.png");
        Picture picture = worksheet.Pictures[pictureIndex];

        // Place the picture inside the cell and anchor it to the merged range
        picture.IsPlacedInCell = true;               // embed the picture in the cell
        picture.MoveToRange(1, 1, 3, 3);              // set the anchor to the merged range

        // Optional: make the picture move and size with the cells (if Placement property is available)
        // picture.Placement = PlacementType.MoveAndSize;

        // Save the workbook
        workbook.Save("PictureInMergedCell.xlsx");
    }
}