using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class VerifyLinkedShapePdf
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing the linked picture shape
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Assume the first picture shape is the linked shape we want to test
            if (sheet.Pictures.Count == 0)
            {
                Console.WriteLine("No picture shapes found in the worksheet.");
                return;
            }

            Picture picture = sheet.Pictures[0];

            // Capture the image bytes of the shape before refresh
            byte[] beforeRefresh;
            using (MemoryStream ms = new MemoryStream())
            {
                picture.ToImage(ms, ImageType.Png);
                beforeRefresh = ms.ToArray();
            }

            // Update the linked cell value that the picture depends on
            // (Replace "B2" with the actual linked cell address if different)
            Cell linkedCell = sheet.Cells["B2"];
            linkedCell.PutValue("NewValue"); // Change the cell value to trigger refresh

            // Refresh the picture so it reflects the new linked cell value
            picture.UpdateSelectedValue();

            // Capture the image bytes of the shape after refresh
            byte[] afterRefresh;
            using (MemoryStream ms = new MemoryStream())
            {
                picture.ToImage(ms, ImageType.Png);
                afterRefresh = ms.ToArray();
            }

            // Simple verification: compare the byte arrays
            bool imagesAreDifferent = !beforeRefresh.SequenceEqual(afterRefresh);
            Console.WriteLine("Image updated after refresh: " + imagesAreDifferent);

            // Save the workbook as PDF – the PDF will contain the refreshed picture
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}