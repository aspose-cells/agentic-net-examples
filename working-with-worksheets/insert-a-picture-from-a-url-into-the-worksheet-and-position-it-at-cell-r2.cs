// Title: How to download an image from a URL and insert it into cell R2 of an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that downloads a PNG image from a given web address and embeds it into cell R2 of the first worksheet using Aspose.Cells. | Generate a .NET example that fetches an image via HttpClient, creates a MemoryStream, and adds the picture to a workbook at row 2, column 18 with Aspose.Cells.
// Common Searches: how to use Aspose.Cells to place an online picture at a specific cell in C# | C# code example for loading an image from the internet and positioning it in Excel with Aspose.Cells | Aspose.Cells tutorial for anchoring a web‑downloaded image to cell R2 | programmatically add a remote PNG to an Excel worksheet using Aspose.Cells .NET
// Tags: aspacells picture insertion from url | c# aspacells anchor image to cell | stream based image embedding in Aspose.Cells | excel workbook save with embedded picture c# | aspacells add picture to worksheet cell

using System;
using System.IO;
using System.Net;
using Aspose.Cells;

// The sample creates a new Workbook, downloads an image from a web URL, adds the picture anchored at cell R2 of the first worksheet, and saves the file as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // URL of the picture to insert
            string imageUrl = "https://example.com/image.png";

            byte[] imageData = null;

            // Attempt to download the image data
            try
            {
                using (WebClient client = new WebClient())
                {
                    imageData = client.DownloadData(imageUrl);
                }
            }
            catch (WebException ex)
            {
                Console.WriteLine($"Failed to download image from '{imageUrl}': {ex.Message}");
                // Continue without inserting the picture
            }

            // Insert the picture if the download succeeded
            if (imageData != null && imageData.Length > 0)
            {
                using (MemoryStream imageStream = new MemoryStream(imageData))
                {
                    // Insert the picture anchored at cell R2 (row index 1, column index 17)
                    sheet.Pictures.Add(1, 17, imageStream);
                }
            }

            // Save the workbook
            string outputPath = "Result.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
