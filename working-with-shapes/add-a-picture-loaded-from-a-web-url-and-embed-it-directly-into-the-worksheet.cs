using System;
using System.IO;
using System.Net.Http;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // URL of the image to embed
            string imageUrl = "https://example.com/sample.jpg";

            // Download the image into a memory stream
            using (HttpClient httpClient = new HttpClient())
            {
                using (Stream imageStream = GetImageStream(httpClient, imageUrl))
                {
                    if (imageStream != null)
                    {
                        // Add the picture to the worksheet at row 1, column 1 (cell B2)
                        worksheet.Pictures.Add(1, 1, imageStream);
                    }
                    else
                    {
                        Console.WriteLine("Image could not be downloaded. Workbook will be saved without picture.");
                    }
                }
            }

            // Save the workbook with the embedded picture
            string outputPath = "WorkbookWithEmbeddedPicture.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to download image and return a stream, handling errors gracefully
    private static Stream GetImageStream(HttpClient client, string url)
    {
        try
        {
            HttpResponseMessage response = client.GetAsync(url).Result;
            response.EnsureSuccessStatusCode();
            return response.Content.ReadAsStreamAsync().Result;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to download image: {ex.Message}");
            return null;
        }
    }
}