using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;
using Aspose.Cells.Rendering;   // Required for ImageOrPrintOptions

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape with specified size and position
            Shape shape = sheet.Shapes.AddAutoShape(
                AutoShapeType.Rectangle, // shape type
                4,   // upper left row
                4,   // top (pixels)
                4,   // upper left column
                4,   // left (pixels)
                200, // height (pixels)
                100  // width (pixels)
            );

            // Set the shape's text
            shape.Text = "Sample Text";

            // Access the text options for the shape's characters
            FontSetting fontSetting = shape.Characters(0, "Sample Text".Length);
            TextOptions textOptions = fontSetting.TextOptions;

            // Apply an outer bottom shadow (preset type OffsetBottom)
            textOptions.Shadow.PresetType = PresetShadowType.OffsetBottom;

            // Capture a screenshot of the shape and save it as an image file
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions(); // defaults to PNG
            string imagePath = "shape.png";
            shape.ToImage(imagePath, imgOptions);
            Console.WriteLine($"Shape image saved to: {Path.GetFullPath(imagePath)}");

            // Save the workbook to a file
            string workbookPath = "output.xlsx";
            workbook.Save(workbookPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(workbookPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}