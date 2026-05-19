using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class ResizeShapeProportionally
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Verify the image file exists before adding it
            string imagePath = "example.jpg";
            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file not found: {imagePath}");

            // Add a picture shape using a FileStream (compatible with all Aspose.Cells versions)
            Shape shape;
            using (FileStream fs = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                shape = worksheet.Shapes.AddPicture(0, 0, 0, 0, fs);
            }

            // Store original dimensions (in points)
            double originalWidth = shape.Width;
            double originalHeight = shape.Height;

            // Define the target range (B2:D5)
            int topRow = 1;      // B2 -> row index 1
            int leftColumn = 1;  // B2 -> column index 1
            int bottomRow = 4;   // D5 -> row index 4
            int rightColumn = 3; // D5 -> column index 3

            // Calculate total width of the target range (pixels)
            double targetWidthPixels = 0;
            for (int col = leftColumn; col <= rightColumn; col++)
                targetWidthPixels += worksheet.Cells.GetColumnWidthPixel(col);

            // Calculate total height of the target range (pixels)
            double targetHeightPixels = 0;
            for (int row = topRow; row <= bottomRow; row++)
                targetHeightPixels += worksheet.Cells.GetRowHeightPixel(row);

            // Convert pixels to points (1 point = 1/72 inch, 1 pixel = 72/96 points)
            const double pixelToPoint = 72.0 / 96.0;
            double targetWidth = targetWidthPixels * pixelToPoint;
            double targetHeight = targetHeightPixels * pixelToPoint;

            // Determine scaling factor while preserving aspect ratio
            double widthScale = targetWidth / originalWidth;
            double heightScale = targetHeight / originalHeight;
            double scale = Math.Min(widthScale, heightScale);

            // Apply proportional resizing (cast to int because Width/Height are int in this API version)
            shape.Width = (int)(originalWidth * scale);
            shape.Height = (int)(originalHeight * scale);

            // Lock aspect ratio (use the updated property)
            shape.IsAspectRatioLocked = true;

            // Move the shape to the target range
            shape.MoveToRange(topRow, leftColumn, bottomRow, rightColumn);

            // Save the workbook
            string outputPath = "ResizeShapeProportionally.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}