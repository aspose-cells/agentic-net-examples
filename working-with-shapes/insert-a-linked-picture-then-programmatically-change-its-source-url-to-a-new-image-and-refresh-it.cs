// Title: Replace a linked picture in an Excel worksheet with a new image URL using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells replace picture with new image from URL | How to update an embedded picture source in an Excel file using Aspose.Cells | Programmatically change picture URL in Aspose.Cells workbook | Aspose.Cells download image stream and refresh picture in worksheet | Remove and add picture at same cell location Aspose.Cells C#
// Tags: replace picture runtime Aspose.Cells | update picture source C# Aspose.Cells | download image stream for Excel picture Aspose.Cells | add picture to worksheet cell Aspose.Cells | remove and reinsert picture Aspose.Cells

using System;
using System.IO;
using System.Net;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The code creates a new workbook, downloads an image from a URL, inserts it as a picture on the first worksheet, removes that picture, inserts a different image from another URL at the same cell location, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // URLs of the images
            string firstImageUrl = "https://example.com/oldimage.png";
            string secondImageUrl = "https://example.com/newimage.png";

            // Row and column where the picture will be placed
            int pictureRow = 1;
            int pictureColumn = 1;

            // Helper to download image data
            byte[] DownloadImage(string url)
            {
                using (WebClient client = new WebClient())
                {
                    return client.DownloadData(url);
                }
            }

            // Add the first picture
            byte[] firstImageData = DownloadImage(firstImageUrl);
            int picIndex;
            using (MemoryStream ms = new MemoryStream(firstImageData))
            {
                picIndex = sheet.Pictures.Add(pictureRow, pictureColumn, ms);
            }

            // Later in the code: replace the picture with a new image
            // Remove the existing picture
            sheet.Pictures.RemoveAt(picIndex);

            // Add the new picture at the same location
            byte[] secondImageData = DownloadImage(secondImageUrl);
            using (MemoryStream ms = new MemoryStream(secondImageData))
            {
                picIndex = sheet.Pictures.Add(pictureRow, pictureColumn, ms);
            }

            // Define output file path
            string outputPath = "LinkedPictureDemo.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
