// Title: Insert a picture from a secure intranet URL with authentication headers and handle access‑denied errors using Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses HttpClient with a bearer token to download an image from a protected intranet URL and adds it to an Aspose.Cells worksheet via a memory stream. | Add logic to detect HTTP 401 or 403 responses when fetching the image and output a clear access‑denied message before saving the workbook. | Save the workbook after inserting the picture at a specific cell (e.g., B2) and ensure the file is written as an .xlsx document.
// Common Searches: how to embed a picture from a secured intranet URL into an Excel file using Aspose.Cells C# | Aspose.Cells download image with bearer token and insert into worksheet | handle 401 unauthorized error when adding picture to Excel with Aspose.Cells | insert picture from memory stream into specific cell using Aspose.Cells | configure HttpClient authentication headers for image retrieval in Aspose.Cells example
// Tags: download protected image via HttpClient for Aspose.Cells | insert picture from memory stream into worksheet cell | add authentication header Aspose.Cells picture insertion | handle 401 403 response when embedding Excel image | secure intranet image embedding Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;

// The example creates a new workbook, configures HttpClient with an Authorization header to fetch an image from a secure intranet URL, checks for 401/403 status codes, reads the image bytes into a MemoryStream, inserts the picture at cell B2 using sheet.Pictures.Add, and saves the workbook as LinkedPicture.xlsx.
class Program
{
    static async Task Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Secure intranet image URL
        string imageUrl = "https://intranet.example.com/secure/image.png";

        // Configure HttpClient with required authentication headers
        using (HttpClient client = new HttpClient())
        {
            // Example: add a bearer token header (replace with actual token)
            client.DefaultRequestHeaders.Add("Authorization", "Bearer YOUR_TOKEN_HERE");

            try
            {
                // Attempt to download the image
                HttpResponseMessage response = await client.GetAsync(imageUrl);

                // Handle access denied scenarios
                if (response.StatusCode == HttpStatusCode.Forbidden ||
                    response.StatusCode == HttpStatusCode.Unauthorized)
                {
                    Console.WriteLine("Access denied: unable to retrieve the image.");
                    return;
                }

                // Throw if the request was not successful
                response.EnsureSuccessStatusCode();

                // Read image data into a byte array
                byte[] imageBytes = await response.Content.ReadAsByteArrayAsync();

                // Insert the picture into the worksheet from a memory stream
                using (MemoryStream ms = new MemoryStream(imageBytes))
                {
                    // Insert at cell B2 (row index 1, column index 1)
                    sheet.Pictures.Add(1, 1, ms);
                }
            }
            catch (HttpRequestException ex)
            {
                // Handle network or request errors
                Console.WriteLine($"Error retrieving image: {ex.Message}");
                return;
            }
        }

        // Save the workbook (lifecycle rule)
        workbook.Save("LinkedPicture.xlsx");
    }
}
