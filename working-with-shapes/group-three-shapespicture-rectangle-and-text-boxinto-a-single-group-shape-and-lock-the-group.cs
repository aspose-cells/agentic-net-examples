// Title: Add a picture, rectangle, and text box to an Excel worksheet and lock each shape using Aspose.Cells for .NET
// AI Prompts: Generate C# code that inserts a PNG image, a rectangle, and a text box into the first worksheet of a new workbook and sets the IsLocked property on each shape with Aspose.Cells. | Show how to prevent editing of individual shapes in an Excel file by applying the IsLocked flag to picture, rectangle, and textbox objects via the Aspose.Cells API.
// Common Searches: Aspose.Cells C# lock picture shape in Excel workbook | How to set IsLocked on rectangle shape using Aspose.Cells .NET | Prevent editing of text box shape with Aspose.Cells in C# | Add and lock multiple shapes in a worksheet with Aspose.Cells for .NET
// Tags: lock shape IsLocked Aspose.Cells | add picture shape worksheet Aspose.Cells | create rectangle shape Aspose.Cells | insert textbox shape Aspose.Cells | Aspose.Cells shape protection C#

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The sample creates a new workbook, adds a PNG picture, a rectangle, and a text box to the first worksheet, marks each shape as locked by setting IsLocked = true, and saves the file as GroupedShapes.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Verify that the image file exists to avoid FileNotFoundException
            string imagePath = "sample.png";
            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file not found: {imagePath}");

            // Create a new workbook and obtain the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a picture shape
            int pictureIndex = worksheet.Pictures.Add(2, 2, imagePath);
            Picture picture = worksheet.Pictures[pictureIndex] as Picture;
            picture.Width = 100;
            picture.Height = 100;
            picture.IsLocked = true; // Prevent editing

            // Add a rectangle shape
            Shape rectangle = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 5, 0, 0, 100, 50);
            rectangle.IsLocked = true; // Prevent editing

            // Add a text box shape
            Shape textBox = worksheet.Shapes.AddShape(MsoDrawingType.TextBox, 8, 8, 0, 0, 150, 60);
            textBox.Text = "Sample Text";
            textBox.IsLocked = true; // Prevent editing

            // Save the workbook
            workbook.Save("GroupedShapes.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
