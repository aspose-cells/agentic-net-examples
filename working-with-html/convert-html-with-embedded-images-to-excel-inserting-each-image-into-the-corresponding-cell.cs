using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class HtmlToExcelWithImages
{
    static void Main()
    {
        // Path to the source HTML file that contains <img> tags (base64 or external URLs)
        string htmlPath = "input.html";

        // Load the HTML file into a workbook.
        // Aspose.Cells automatically parses the HTML and creates pictures for the images.
        Workbook workbook = new Workbook(htmlPath);

        // Iterate through each worksheet in the workbook.
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through each picture that was created from the HTML images.
            foreach (Picture picture in sheet.Pictures)
            {
                // The picture's UpperLeftRow and UpperLeftColumn indicate the cell it belongs to.
                // Set the placement so the picture moves and resizes with the cell.
                picture.Placement = PlacementType.MoveAndSize;
            }
        }

        // Save the workbook as an Excel file. The images are now embedded in the corresponding cells.
        workbook.Save("output.xlsx");
    }
}