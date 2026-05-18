using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class RefreshLinkedPicturesParallel
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // List of external image URLs
        List<string> imageUrls = new List<string>
        {
            "https://example.com/image1.jpg",
            "https://example.com/image2.jpg",
            "https://example.com/image3.jpg",
            // Add more URLs as needed
        };

        // Add linked pictures in parallel
        Parallel.ForEach(imageUrls, (url, state, index) =>
        {
            // Simple grid positioning
            int row = (int)(index / 3) * 15;   // 15 rows spacing
            int col = (int)(index % 3) * 15;   // 15 columns spacing

            // Add a linked picture (height and width in pixels)
            sheet.Shapes.AddLinkedPicture(row, col, 100, 100, url);
        });

        // Save the workbook
        workbook.Save("LinkedPicturesParallel.xlsx");
    }
}