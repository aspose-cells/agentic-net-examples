using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Path to the image that will be used as a footer watermark
        string imagePath = "footer_watermark.png";

        // Verify that the image file exists
        if (!File.Exists(imagePath))
        {
            Console.WriteLine($"Image file not found: {imagePath}");
            return;
        }

        // Load image data into a byte array
        byte[] imageData = File.ReadAllBytes(imagePath);

        // Create a new workbook (you can also load an existing one if needed)
        Workbook workbook = new Workbook();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Set the picture in the center section of the footer (section index 1)
            sheet.PageSetup.SetFooterPicture(1, imageData);

            // Set the footer script to display the picture (&G)
            sheet.PageSetup.SetFooter(1, "&G");
        }

        // Save the workbook to a file
        string outputPath = "WorkbookWithFooterWatermark.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved successfully to {outputPath}");
    }
}