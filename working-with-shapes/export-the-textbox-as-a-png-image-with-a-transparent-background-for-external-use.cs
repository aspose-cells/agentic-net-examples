using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class ExportTextboxToTransparentPng
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a textbox shape to the worksheet
        // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, width, height
        Shape textbox = sheet.Shapes.AddTextBox(2, 1, 0, 0, 200, 80);
        textbox.Text = "Transparent TextBox";

        // Prepare image options: PNG format with transparent background
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,
            Transparent = true
        };

        // Export the textbox to a PNG file using the options above
        string outputPath = "textbox_transparent.png";
        textbox.ToImage(outputPath, imgOptions);

        Console.WriteLine($"Textbox exported to PNG with transparent background: {outputPath}");
    }
}