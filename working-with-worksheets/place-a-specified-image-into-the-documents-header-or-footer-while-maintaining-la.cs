using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class HeaderFooterImageDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        PageSetup pageSetup = worksheet.PageSetup;

        // Load the image file into a byte array; if not found, skip adding the image
        string imagePath = Path.Combine(Environment.CurrentDirectory, "logo.png");
        if (File.Exists(imagePath))
        {
            byte[] imageData = File.ReadAllBytes(imagePath);

            // Set the image in the center section of the header
            pageSetup.SetHeaderPicture(1, imageData);
            pageSetup.SetHeader(1, "&G");

            // Set the same image in the right section of the footer
            pageSetup.SetFooterPicture(2, imageData);
            pageSetup.SetFooter(2, "&G");
        }

        // Save the workbook with the header/footer images
        workbook.Save("HeaderFooterImage.xlsx");
    }
}