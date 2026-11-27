// Title: Insert an image from a web URL into cell R2 of an Excel worksheet using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells download image from URL and add to worksheet cell | How to place a picture at R2 in Excel with Aspose.Cells .NET | Insert picture into specific cell using Aspose.Cells Pictures.Add method | Set picture name after inserting image in Aspose.Cells workbook
// Tags: Aspose.Cells picture insertion from URL | Worksheet.Pictures.Add cell positioning | Excel image placement at specific cell .NET | download image stream for Aspose.Cells | assign picture name Aspose.Cells workbook

using Aspose.Cells;
using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, downloads an image from a given URL using HttpClient, inserts the image into the first worksheet at cell R2 (row 1, column 17) with Worksheet.Pictures.Add, optionally sets a name for the picture, and saves the workbook as Output.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // URL of the picture to insert
                string imageUrl = "https://example.com/image.png";

                // Download the image and insert it if the download succeeds
                InsertPictureFromUrl(sheet, imageUrl).Wait();

                // Save the workbook to a file
                string outputPath = "Output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private static async Task InsertPictureFromUrl(Worksheet sheet, string url)
        {
            try
            {
                using (HttpClient client = new HttpClient())
                using (HttpResponseMessage response = await client.GetAsync(url))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        Console.WriteLine($"Failed to download image. HTTP {(int)response.StatusCode} - {response.ReasonPhrase}");
                        return;
                    }

                    using (Stream imageStream = await response.Content.ReadAsStreamAsync())
                    {
                        // Add the picture at cell R2 (row index 1, column index 17)
                        int pictureIndex = sheet.Pictures.Add(1, 17, imageStream);
                        // Optional: set a name for the picture
                        sheet.Pictures[pictureIndex].Name = "InsertedPicture";
                        Console.WriteLine("Picture inserted successfully.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception while inserting picture: {ex.Message}");
            }
        }
    }
}
