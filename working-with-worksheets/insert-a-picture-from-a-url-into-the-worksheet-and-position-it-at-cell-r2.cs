// Title: Download an image from a URL and insert it into cell R2 of an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses HttpClient to fetch an image from a given URL, then adds the image to cell R2 of the first worksheet in an Aspose.Cells workbook and saves the file. | Generate a C# example that creates a new Workbook, streams a PNG downloaded via HttpClient into a MemoryStream, inserts it at row 2 column R with Worksheet.Pictures.Add, and writes the result to Result.xlsx. | Provide a C# snippet that catches HttpRequestException while downloading an image and falls back to a local placeholder before inserting the picture into cell R2 with Aspose.Cells.
// Common Searches: Aspose.Cells C# insert image from internet into specific cell R2 | download picture with HttpClient and add to Excel worksheet using Aspose.Cells | place external PNG at cell R2 in .NET Excel file Aspose.Cells example | how to use Worksheet.Pictures.Add with a MemoryStream in C#
// Tags: Aspose.Cells download image stream insertion | Worksheet.Pictures.Add from HttpClient stream | C# insert web image into Excel cell | Aspose.Cells place picture at specific cell | Excel workbook save after picture insertion .NET

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample creates a new workbook, downloads an image from a web URL using HttpClient, inserts the image into cell R2 of the first worksheet via Worksheet.Pictures.Add, and saves the workbook as Result.xlsx, with error handling for download failures.
    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // URL of the picture to insert
                string imageUrl = "https://example.com/image.png";

                // Download the image and insert it into the worksheet
                using (HttpClient client = new HttpClient())
                {
                    try
                    {
                        using (Stream downloadStream = await client.GetStreamAsync(imageUrl))
                        using (MemoryStream pictureStream = new MemoryStream())
                        {
                            await downloadStream.CopyToAsync(pictureStream);
                            pictureStream.Position = 0; // Reset stream position

                            // Insert the picture at cell R2 (row index 1, column index 17)
                            int row = 1;      // Zero‑based row index for row 2
                            int column = 17;  // Zero‑based column index for column R
                            sheet.Pictures.Add(row, column, pictureStream);
                        }
                    }
                    catch (HttpRequestException ex)
                    {
                        Console.WriteLine($"Image download failed: {ex.Message}");
                        // If needed, handle fallback logic here (e.g., use a local placeholder image)
                    }
                }

                // Save the workbook to a file
                string outputPath = "Result.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
