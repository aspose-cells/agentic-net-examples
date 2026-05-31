using System;
using System.IO;
using Aspose.Cells;

class EmbedImageInMergedCell
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells B2 and C2 (row index 1, column index 1, 1 row, 2 columns)
        worksheet.Cells.Merge(1, 1, 1, 2);

        // Load an image file into a byte array
        string imagePath = "sample.png";               // replace with your image file path
        byte[] imageBytes = File.ReadAllBytes(imagePath);

        // Convert the image bytes to a Base64 string
        string base64Image = Convert.ToBase64String(imageBytes);

        // Build an HTML img tag that uses the Base64 data URI
        string htmlImgTag = $"<img src=\"data:image/png;base64,{base64Image}\" style=\"width:100%;height:auto;\"/>";

        // Set the HtmlString of the top‑left cell of the merged range (B2)
        worksheet.Cells["B2"].HtmlString = htmlImgTag;

        // Save the workbook; the image will be rendered inside the merged cell when opened
        workbook.Save("MergedCellWithImage.xlsx");
    }
}