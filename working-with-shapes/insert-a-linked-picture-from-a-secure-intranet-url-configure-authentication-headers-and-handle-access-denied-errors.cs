using System;
using System.IO;
using System.Net;
using System.Net.Http;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class LinkedPictureExample
{
    static void Main()
    {
        // URL of the image on the secure intranet
        string imageUrl = "https://intranet.example.com/secureimage.jpg";

        // Create an HttpClient to download the image with authentication headers
        using (HttpClient client = new HttpClient())
        {
            // Example: add a bearer token; replace with actual authentication method
            client.DefaultRequestHeaders.Add("Authorization", "Bearer YOUR_ACCESS_TOKEN");

            try
            {
                // Send request synchronously
                HttpResponseMessage response = client.GetAsync(imageUrl).Result;

                // Handle access denied (HTTP 403) explicitly
                if (response.StatusCode == HttpStatusCode.Forbidden)
                {
                    Console.WriteLine("Access denied while retrieving the image. Check credentials or permissions.");
                    return;
                }

                // Throw if other non-success status codes are returned
                response.EnsureSuccessStatusCode();

                // Save the downloaded image to a temporary file
                string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + Path.GetExtension(imageUrl));
                using (FileStream fs = new FileStream(tempFilePath, FileMode.Create, FileAccess.Write))
                {
                    response.Content.CopyToAsync(fs).Wait();
                }

                // Create a new workbook and add the linked picture
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add the linked picture using the ShapeCollection.AddLinkedPicture method (rule)
                // Parameters: topRow, leftColumn, height (px), width (px), sourceFullName (file path)
                worksheet.Shapes.AddLinkedPicture(1, 1, 200, 200, tempFilePath);

                // Save the workbook
                string outputPath = "LinkedPicture.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (HttpRequestException ex)
            {
                // Handle network errors, DNS failures, etc.
                Console.WriteLine($"Error downloading image: {ex.Message}");
            }
            catch (Exception ex)
            {
                // General exception handling
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}