// Title: Embedding a Base64‑encoded PNG into a merged Excel range using the HtmlString property with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# insert image into merged cells using HTML | display a PNG in a merged Excel range with HtmlString property | C# embed base64 image in Excel cell with Aspose.Cells | how to set row height and column width for images in merged cells Aspose.Cells | load image file and convert to data URI for Excel using Aspose.Cells .NET
// Tags: HtmlString image embedding Aspose.Cells | merged range picture rendering Aspose.Cells | Base64 data URI for Excel cell image | C# image to Base64 conversion Aspose.Cells | row height column width adjustment for embedded images

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, merges cells B2:D4, reads a PNG file, converts it to a Base64 data URI, and assigns an <img> tag to the merged cell via the HtmlString property. Row heights and column widths are adjusted for proper image display, and the workbook is saved as output.xlsx, with fallback text if the image is missing.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells B2:D4 (rows 1‑3, columns 1‑3)
            sheet.Cells.Merge(1, 1, 3, 3);

            // Path to the image file
            string imagePath = "sample.png";
            string imgSrc = string.Empty;

            if (File.Exists(imagePath))
            {
                // Load image bytes and convert to Base64 data URI
                byte[] imageBytes = File.ReadAllBytes(imagePath);
                string base64 = Convert.ToBase64String(imageBytes);
                imgSrc = $"data:image/png;base64,{base64}";

                // Insert HTML img tag into the merged cell (top‑left cell B2)
                Cell targetCell = sheet.Cells[1, 1];
                targetCell.HtmlString = $"<img src=\"{imgSrc}\" style=\"width:100%;height:auto;\"/>";
            }
            else
            {
                // If the image is missing, place a placeholder text
                sheet.Cells[1, 1].PutValue("Image not found");
            }

            // Adjust row heights and column widths for better display
            for (int r = 1; r <= 3; r++) sheet.Cells.SetRowHeight(r, 80);
            for (int c = 1; c <= 3; c++) sheet.Cells.SetColumnWidth(c, 20);

            // Save the workbook
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
