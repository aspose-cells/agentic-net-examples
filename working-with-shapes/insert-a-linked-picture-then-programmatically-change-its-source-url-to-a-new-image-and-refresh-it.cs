// Title: Insert a linked picture into an Excel worksheet and programmatically change its source URL with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add picture linked to external URL in Excel | How to update the URL of a linked image in an existing workbook using Aspose.Cells | Replace placeholder PNG with web image programmatically in Aspose.Cells | Refresh linked picture source after changing URL in Aspose.Cells .NET | Change linked picture URL without reopening Excel using Aspose.Cells API
// Tags: add linked picture Aspose.Cells | update picture URL Aspose.Cells | replace placeholder image with web URL Aspose.Cells | linked image refresh Excel .NET | programmatic picture source change Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a tiny transparent PNG placeholder, inserts it as a picture, removes it, adds a linked picture from an initial URL, then replaces that linked picture with a new URL, and finally saves the workbook as 'LinkedPictureDemo.xlsx' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Path for a tiny transparent placeholder image.
            const string placeholderPath = "placeholder.png";

            // Create the placeholder file if it does not exist.
            if (!File.Exists(placeholderPath))
            {
                // 1x1 transparent PNG (base64 encoded).
                const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+Z3WcAAAAASUVORK5CYII=";
                byte[] pngBytes = Convert.FromBase64String(base64Png);
                File.WriteAllBytes(placeholderPath, pngBytes);
            }

            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add the placeholder picture (returns the picture index).
            int picIndex = sheet.Pictures.Add(1, 1, placeholderPath);
            Aspose.Cells.Drawing.Picture picture = sheet.Pictures[picIndex];

            // Replace the placeholder with an image from a URL.
            // Remove the placeholder picture.
            sheet.Pictures.RemoveAt(picIndex);

            // Add a picture linked to the first URL.
            int linkedPicIdx = sheet.Pictures.Add(1, 1, "https://example.com/oldimage.png");
            // (Aspose.Cells automatically handles linked images when a URL is supplied.)

            // Replace the linked picture with a new URL.
            sheet.Pictures.RemoveAt(linkedPicIdx);
            sheet.Pictures.Add(1, 1, "https://example.com/newimage.png");

            // Save the workbook.
            workbook.Save("LinkedPictureDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
