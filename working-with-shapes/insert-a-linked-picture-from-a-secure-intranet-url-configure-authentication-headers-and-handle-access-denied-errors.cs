// Title: Embed a Secured Intranet Image into Excel with Aspose.Cells (.NET) – Auth Headers & 403 Handling
// Description: Download a JPEG from a protected intranet URL using HttpClient with Bearer and custom headers, detect 403 Forbidden, embed the image into a worksheet via a MemoryStream, and save the workbook with Aspose.Cells for .NET.
// Keywords: Aspose.Cells insert picture | C# download image with bearer token | secure intranet image Excel | handle 403 forbidden Aspose.Cells | embed image worksheet Aspose.Cells | HttpClient authentication headers | linked picture vs embedded Aspose.Cells
// Common Searches: How to add a picture from a protected intranet URL using Aspose.Cells .NET | Aspose.Cells download image with bearer token | Insert image into Excel workbook with custom HTTP headers | Handle 403 Forbidden when embedding picture in Aspose.Cells | Embed versus link picture in Aspose.Cells
// Developer Intent: Retrieve an image from a secured intranet endpoint with required authentication, embed it into an Excel worksheet, and gracefully handle access‑denied responses.
// Use Cases: Use HttpClient to request the image, add Authorization and any custom headers, and check for a 403 status before proceeding. | Create a MemoryStream from the downloaded byte array and call worksheet.Pictures.Add(row, column, stream) to place the picture at a specific cell. | Save the workbook after embedding the picture, producing an .xlsx file that contains the secured image.
// AI Prompts: Generate C# code that uses Aspose.Cells to embed a picture from a secure intranet URL with a Bearer token and custom headers, including 403 error handling. | Show an example of downloading an image via HttpClient, adding authentication headers, and inserting it into an Excel worksheet with Aspose.Cells. | Explain how to switch from an embedded picture to a linked picture in Aspose.Cells while preserving authentication requirements.

using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Download a JPEG from a protected intranet URL using HttpClient with Bearer and custom headers, detect 403 Forbidden, embed the image into a worksheet via a MemoryStream, and save the workbook with Aspose.Cells for .NET.
class Program
{
    static async Task Main()
    {
        // Secure intranet image URL
        string imageUrl = "https://intranet.example.com/secure/image.jpg";

        // Configure HttpClient with required authentication headers
        using var httpClient = new HttpClient();
        httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer YOUR_ACCESS_TOKEN");
        httpClient.DefaultRequestHeaders.Add("Custom-Header", "CustomValue");

        byte[] imageBytes;

        try
        {
            // Attempt to download the image
            HttpResponseMessage response = await httpClient.GetAsync(imageUrl);

            // Handle access denied (403) explicitly
            if (response.StatusCode == HttpStatusCode.Forbidden)
            {
                Console.WriteLine("Access denied: unable to retrieve the image from the intranet URL.");
                return;
            }

            response.EnsureSuccessStatusCode(); // Throw if not successful
            imageBytes = await response.Content.ReadAsByteArrayAsync();
        }
        catch (Exception ex)
        {
            // General error handling for network issues, timeouts, etc.
            Console.WriteLine($"Error retrieving image: {ex.Message}");
            return;
        }

        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert the image into the worksheet.
        // Since linked pictures cannot carry custom HTTP headers, we embed the downloaded image.
        using (MemoryStream ms = new MemoryStream(imageBytes))
        {
            // Add picture at row 2, column 2 (zero‑based indices)
            worksheet.Pictures.Add(2, 2, ms);
        }

        // Save the workbook (lifecycle rule)
        workbook.Save("IntranetLinkedPicture.xlsx");
        Console.WriteLine("Workbook saved with the image embedded.");
    }
}
