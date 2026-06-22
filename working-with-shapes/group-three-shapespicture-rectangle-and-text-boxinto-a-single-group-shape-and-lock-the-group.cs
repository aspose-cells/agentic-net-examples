using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsDemo
{
    public class GroupShapesDemo
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
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                ShapeCollection shapes = sheet.Shapes;

                // Path to the image file
                string imagePath = "image.png";

                // Verify that the image file exists before adding it
                if (!File.Exists(imagePath))
                    throw new FileNotFoundException("Image file not found.", imagePath);

                // Add a picture shape using a FileStream (required overload)
                Shape picture;
                using (FileStream imgStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                {
                    picture = shapes.AddPicture(1, 0, 100, 100, imgStream);
                }

                // Add a rectangle shape (upper left row, column, lower right row, column, height, width)
                Shape rectangle = shapes.AddRectangle(3, 0, 3, 0, 120, 80);

                // Add a text box shape and set its text
                Shape textBox = shapes.AddTextBox(5, 0, 5, 0, 150, 60);
                textBox.Text = "Sample Text";

                // Group the three shapes
                Shape[] itemsToGroup = new Shape[] { picture, rectangle, textBox };
                GroupShape group = shapes.Group(itemsToGroup);

                // Lock the group shape
                group.IsLocked = true;

                // Save the workbook
                string outputPath = "GroupShapesLocked.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Runtime safety: report any errors that occur during processing
                Console.WriteLine($"Run error: {ex.Message}");
            }
        }
    }
}