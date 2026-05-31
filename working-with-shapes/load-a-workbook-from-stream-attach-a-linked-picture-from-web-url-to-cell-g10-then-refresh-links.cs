using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Load the workbook from a stream (replace with your actual stream source)
        using (FileStream inputStream = new FileStream("input.xlsx", FileMode.Open, FileAccess.Read))
        {
            Workbook workbook = new Workbook(inputStream);
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a linked picture to cell G10 (zero‑based row 9, column 6)
            string imageUrl = "https://example.com/sample.jpg";
            int pictureHeight = 100; // pixels
            int pictureWidth  = 100; // pixels
            worksheet.Shapes.AddLinkedPicture(9, 6, pictureHeight, pictureWidth, imageUrl);

            // Refresh external links (if any exist)
            workbook.UpdateLinkedDataSource(new Workbook[] { });

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}