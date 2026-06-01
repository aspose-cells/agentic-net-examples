using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsRowCopyDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate source rows with sample data
                cells["A1"].PutValue("Header");
                cells["A2"].PutValue("Row 1 Data");
                cells["A3"].PutValue("Row 2 Data");

                // Insert a picture that resides in row 2 (index 1) if the file exists
                string picturePath = "sample_image.png";
                if (File.Exists(picturePath))
                {
                    int pictureIndex = sheet.Pictures.Add(1, 0, picturePath); // row 1, column 0
                    Picture picture = sheet.Pictures[pictureIndex];
                    picture.Width = 100;
                    picture.Height = 50;
                }
                else
                {
                    Console.WriteLine($"Warning: Picture file '{picturePath}' not found. Skipping picture insertion.");
                }

                // Insert a rectangle shape that also resides in row 2
                ShapeCollection shapes = sheet.Shapes;
                // topRow, leftColumn, bottomRow, rightColumn, width, height
                RectangleShape rect = shapes.AddRectangle(1, 0, 1, 0, 120, 60);

                // Define source and destination row indices
                int sourceStartRow = 0;          // start copying from row 0 (Header)
                int destinationStartRow = 5;    // paste starting at row 5
                int rowsToCopy = 3;             // copy three rows (0,1,2)

                // Perform the copy (preserves images and drawing objects)
                cells.CopyRows(cells, sourceStartRow, destinationStartRow, rowsToCopy);

                // Save the workbook
                string outputPath = "RowCopyWithImagesAndShapes.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}