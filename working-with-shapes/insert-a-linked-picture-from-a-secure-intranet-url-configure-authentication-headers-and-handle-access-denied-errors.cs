// Title: Add a Linked Picture from a Secure Intranet URL with Authentication Headers and Fallback Embedding – Aspose.Cells for .NET
// Description: Demonstrates how to insert a linked picture from a protected intranet URL into an Excel worksheet using Aspose.Cells. The example shows how to supply an Authorization header with HttpClient, catch access‑denied errors from Shapes.AddLinkedPicture, download the image manually, and embed it from a stream as a fallback before saving the workbook.
// Keywords: Aspose.Cells linked picture authentication | C# add linked picture from secure URL | HttpClient Authorization header Excel | fallback embed image Aspose.Cells | handle access denied linked picture | intranet image Excel workbook | .NET Aspose.Cells Shapes.AddLinkedPicture
// Common Searches: Aspose.Cells add linked picture with bearer token | C# insert picture from protected URL in Excel | How to handle access denied when adding linked picture Aspose.Cells | Download image with HttpClient for Aspose.Cells | Fallback to embed picture if linked picture fails
// Developer Intent: Insert a picture from a secured intranet URL into an Excel sheet, providing authentication headers and automatically falling back to embedding the image when the linked picture cannot be accessed.
// Use Cases: Corporate dashboards that reference images stored behind an authentication gateway. | Automated report generation where the source image requires an OAuth2 bearer token. | Scenarios where network policies block external linking, requiring the image to be embedded locally. | Logging HTTP status codes for failed image retrieval while still producing a valid workbook.
// AI Prompts: Write C# code using Aspose.Cells to add a linked picture from a URL that needs a Bearer token, with error handling that falls back to embedding the image from a stream. | Explain how to configure HttpClient default request headers for authentication when downloading an image for Aspose.Cells, and how to detect and handle 401/403 responses. | Provide a step‑by‑step guide to test linked picture insertion against a mock secure server and verify that the fallback embedding works correctly.

using System;
using System.IO;
using System.Net;
using System.Net.Http;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to insert a linked picture from a protected intranet URL into an Excel worksheet using Aspose.Cells. The example shows how to supply an Authorization header with HttpClient, catch access‑denied errors from Shapes.AddLinkedPicture, download the image manually, and embed it from a stream as a fallback before saving the workbook.
class InsertLinkedPictureWithAuth
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Secure intranet image URL
            string imageUrl = "https://intranet.example.com/secure/image.jpg";

            // Authentication header (example uses Bearer token)
            const string authHeaderName = "Authorization";
            const string authHeaderValue = "Bearer YOUR_ACCESS_TOKEN";

            // Try to add a linked picture directly
            try
            {
                // AddLinkedPicture uses row/column indices (0‑based) and pixel dimensions
                worksheet.Shapes.AddLinkedPicture(2, 2, 150, 150, imageUrl);
            }
            catch (Exception ex) when (ex is WebException || ex is CellsException)
            {
                // If the linked picture cannot be created (e.g., access denied), download the image manually
                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add(authHeaderName, authHeaderValue);
                    HttpResponseMessage response = client.GetAsync(imageUrl).Result;

                    if (response.IsSuccessStatusCode)
                    {
                        using (Stream imageStream = response.Content.ReadAsStreamAsync().Result)
                        {
                            // Embed the picture from the downloaded stream as a fallback
                            worksheet.Pictures.Add(2, 2, imageStream);
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Failed to retrieve image. HTTP status: {response.StatusCode}");
                    }
                }
            }

            // Save the workbook
            string outputPath = "LinkedPictureDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception e)
        {
            Console.WriteLine($"Unexpected error: {e.Message}");
        }
    }
}
