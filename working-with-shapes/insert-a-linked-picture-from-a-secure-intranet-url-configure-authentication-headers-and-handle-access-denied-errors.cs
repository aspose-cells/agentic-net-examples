// Title: Insert a picture from a secured intranet URL into an Excel worksheet with Aspose.Cells for .NET, using HttpClient authentication and handling 403 errors
// Common Searches: Aspose.Cells C# insert image from protected intranet URL with bearer token | How to handle HTTP 403 when adding a picture to an Excel file using Aspose.Cells | Load picture into Aspose.Cells worksheet from HttpClient stream | Add external image to Excel using Aspose.Cells with custom authentication headers | C# Aspose.Cells picture insertion from secure web service
// Tags: insert picture from stream Aspose.Cells C# | authenticated image retrieval Aspose.Cells | handle HTTP 403 Aspose.Cells picture insertion | secure intranet image Aspose.Cells | Aspose.Cells picture add with HttpClient

using Aspose.Cells;
using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;

// The example creates a new Workbook, configures an HttpClient with an Authorization header, fetches an image from a secure intranet URL, checks for a 403 Forbidden response, and inserts the image into the first worksheet as a picture from a memory stream, setting its size. On errors it writes a message to cell A1, then saves the workbook as LinkedPicture.xlsx.
class Program
{
    static async Task Main()
    {
        // Create a new workbook (lifecycle rule: create)
        var workbook = new Workbook();

        // Access the first worksheet
        var sheet = workbook.Worksheets[0];

        // Secure intranet URL of the picture
        string imageUrl = "https://intranet.example.com/images/logo.png";

        // Configure HttpClient with authentication headers
        using var httpClient = new HttpClient();
        // Example: add a bearer token (replace with actual token or other auth scheme)
        httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer YOUR_ACCESS_TOKEN");

        try
        {
            // Request the image from the intranet
            using var response = await httpClient.GetAsync(imageUrl);

            // Handle access denied (HTTP 403) explicitly
            if (response.StatusCode == HttpStatusCode.Forbidden)
            {
                Console.WriteLine("Access denied to the image URL.");
                sheet.Cells["A1"].PutValue("Image not available (Access Denied)");
            }
            else
            {
                // Throw if the request was not successful
                response.EnsureSuccessStatusCode();

                // Read the image bytes
                var imageBytes = await response.Content.ReadAsByteArrayAsync();

                // Add the picture to the worksheet from a stream
                using var ms = new MemoryStream(imageBytes);
                int pictureIndex = sheet.Pictures.Add(2, 2, ms); // Row 2, Column 2 (zero‑based)
                var picture = sheet.Pictures[pictureIndex];

                // Optional: set picture size
                picture.Width = 200;
                picture.Height = 100;
            }
        }
        catch (HttpRequestException ex)
        {
            // Handle HTTP errors (e.g., network issues)
            Console.WriteLine($"Error retrieving image: {ex.Message}");
            sheet.Cells["A1"].PutValue("Image not available (Error)");
        }
        catch (Exception ex)
        {
            // General exception handling
            Console.WriteLine($"Unexpected error: {ex.Message}");
            sheet.Cells["A1"].PutValue("Image not available (Unexpected Error)");
        }

        try
        {
            // Save the workbook (lifecycle rule: save)
            workbook.Save("LinkedPicture.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
