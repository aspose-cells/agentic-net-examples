using System;
using System.IO;
using Aspose.Cells;

class InsertImageMarker
{
    static void Main()
    {
        // Load the existing workbook that contains the image marker (e.g., ${image}) in a cell
        Workbook workbook = new Workbook("template.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Locate the cell that holds the image marker
        Cell imageMarkerCell = null;
        foreach (Cell cell in worksheet.Cells)
        {
            if (cell.StringValue == "${image}")
            {
                imageMarkerCell = cell;
                break;
            }
        }

        if (imageMarkerCell != null)
        {
            // Load image data into a byte array (replace with your actual image file)
            byte[] imageBytes = File.ReadAllBytes("picture.png");

            // Assign the image bytes to the cell's EmbeddedImage property
            imageMarkerCell.EmbeddedImage = imageBytes;

            // Optionally clear the placeholder text after embedding the image
            imageMarkerCell.PutValue(string.Empty);
        }
        else
        {
            Console.WriteLine("Image marker not found in the worksheet.");
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}