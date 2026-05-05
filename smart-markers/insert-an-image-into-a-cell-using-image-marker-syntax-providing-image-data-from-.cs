using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsImageMarkerDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the template workbook that contains an image marker (e.g., "{{ImageMarker}}")
            string templatePath = "Template.xlsx";

            // Load the workbook (lifecycle rule: use provided load method)
            Workbook workbook = new Workbook(templatePath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Locate the cell that holds the image marker text
            // Adjust the marker text to match the placeholder used in your template
            string imageMarker = "{{ImageMarker}}";
            Cell markerCell = worksheet.Cells.Find(imageMarker, null, new FindOptions());

            if (markerCell == null)
            {
                Console.WriteLine($"Image marker \"{imageMarker}\" not found in the worksheet.");
                return;
            }

            // Example: load image bytes from a file (could also be from a database, API, etc.)
            byte[] imageBytes = File.ReadAllBytes("sample.png");

            // Insert the image into the cell using the EmbeddedImage property
            // This replaces the marker text with the actual image
            markerCell.EmbeddedImage = imageBytes;

            // Optionally clear the cell's text value (the image will be displayed in the cell)
            markerCell.PutValue(string.Empty);

            // Save the modified workbook (lifecycle rule: use provided save method)
            string outputPath = "ResultWithImage.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Image inserted and workbook saved to: {Path.GetFullPath(outputPath)}");
        }
    }
}