// Title: Insert a picture from a protected intranet URL with bearer token authentication and add fallback text shapes for HTTP 403 and other errors using Aspose.Cells for .NET
// AI Prompts: Generate C# code that configures HttpClient with an Authorization header, downloads an image from a secure intranet URL, and embeds it into an Aspose.Cells worksheet, inserting a text shape when the response is 403 Forbidden. | Demonstrate how to catch network and HTTP exceptions while adding a picture to an Excel file with Aspose.Cells and replace the picture with a descriptive text shape for any failure.
// Common Searches: how to add a picture from an authenticated https endpoint to an Aspose.Cells workbook in C# | Aspose.Cells insert image with bearer token and handle 403 forbidden response | C# download image using HttpClient and embed into Excel using Aspose.Cells | fallback to text shape when image download fails in Aspose.Cells worksheet | ignore SSL certificate errors for intranet image download with Aspose.Cells
// Tags: authenticated image download Aspose.Cells | insert picture from secure URL C# | fallback text shape for HTTP error Aspose.Cells | bearer token HttpClient Aspose.Cells | ignore SSL certificate Aspose.Cells intranet

using System;
using System.IO;
using System.Net;
using System.Net.Http;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to use HttpClient with a bearer token to retrieve an image from a protected intranet URL, embed the image into an Aspose.Cells worksheet, and insert descriptive text shapes as placeholders for 403 Forbidden, other HTTP errors, or exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Secure intranet image URL
            var imageUrl = "https://intranet.example.com/secure/image.png";

            // Configure HttpClient with authentication header
            using (var handler = new HttpClientHandler())
            {
                // Accept all SSL certificates for intranet testing (use with caution)
                handler.ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => true;

                using (var client = new HttpClient(handler))
                {
                    // Example: Bearer token authentication
                    client.DefaultRequestHeaders.Add("Authorization", "Bearer YOUR_ACCESS_TOKEN");

                    try
                    {
                        // Request the image
                        var response = client.GetAsync(imageUrl).Result;

                        if (response.IsSuccessStatusCode)
                        {
                            // Load image into a memory stream
                            using (var imageStream = new MemoryStream(response.Content.ReadAsByteArrayAsync().Result))
                            {
                                // Insert picture (embedded). Aspose.Cells does not expose IsLinked/LinkSource in this version.
                                int pictureIndex = sheet.Pictures.Add(1, 1, imageStream);
                                // Picture picture = sheet.Pictures[pictureIndex]; // reference if further manipulation is needed
                            }
                        }
                        else if (response.StatusCode == HttpStatusCode.Forbidden)
                        {
                            // Access denied – insert a placeholder text shape
                            var shape = sheet.Shapes.AddTextEffect(
                                MsoPresetTextEffect.TextEffect1,
                                "Access Denied",
                                "Arial",
                                12,
                                false,
                                false,
                                1,
                                1,
                                200,
                                30,
                                0,   // rotation
                                0);  // transparency
                            shape.Placement = PlacementType.FreeFloating;
                        }
                        else
                        {
                            // Other HTTP errors – insert a generic error placeholder
                            var shape = sheet.Shapes.AddTextEffect(
                                MsoPresetTextEffect.TextEffect1,
                                $"Error: {response.StatusCode}",
                                "Arial",
                                12,
                                false,
                                false,
                                1,
                                1,
                                200,
                                30,
                                0,
                                0);
                            shape.Placement = PlacementType.FreeFloating;
                        }
                    }
                    catch (Exception ex)
                    {
                        // Network or unexpected exception – insert exception message
                        var shape = sheet.Shapes.AddTextEffect(
                            MsoPresetTextEffect.TextEffect1,
                            $"Exception: {ex.Message}",
                            "Arial",
                            12,
                            false,
                            false,
                            1,
                            1,
                            300,
                            30,
                            0,
                            0);
                        shape.Placement = PlacementType.FreeFloating;
                    }
                }
            }

            // Save the workbook
            workbook.Save("LinkedPicture.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
