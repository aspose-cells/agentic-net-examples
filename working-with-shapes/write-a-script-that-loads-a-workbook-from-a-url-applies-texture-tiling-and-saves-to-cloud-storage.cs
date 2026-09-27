// Title: Download an Excel workbook from a URL, set a solid background on a cell range, and upload the updated file to cloud storage using Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to fetch an .xlsx file from a specified HTTP endpoint, apply a light‑gray solid fill to cells A1:B2, and then send the modified workbook to a cloud storage URL with an HTTP PUT request. | Create a .NET console program that loads an Excel workbook into a MemoryStream via HttpClient, styles a range with a solid background using Aspose.Cells, saves the workbook as Xlsx bytes, and uploads the result to a remote storage service.
// Common Searches: Aspose.Cells C# download workbook from web URL, apply background color to range, and upload to Azure Blob storage | How to modify an Excel file in memory with Aspose.Cells and send it to Amazon S3 using HTTP PUT in C# | C# example for applying solid fill to cells A1:B2 with Aspose.Cells after downloading the file from a remote server | Upload a processed .xlsx file to cloud storage directly from a .NET application using Aspose.Cells and HttpClient
// Tags: download Excel workbook with HttpClient Aspose.Cells | apply solid fill to cell range Aspose.Cells | save workbook to memory stream Xlsx | upload .xlsx via HTTP PUT C# | in‑memory Excel manipulation Aspose.Cells .NET

using System;
using System.IO;
using System.Net.Http;
using Aspose.Cells;

// Downloads an .xlsx file from a remote URL, applies a light‑gray solid background to cells A1:B2 using Aspose.Cells, saves the workbook to a memory stream as Xlsx, and uploads the resulting file to a cloud storage endpoint via HTTP PUT.
class Program
{
    static void Main()
    {
        try
        {
            // URL of the source workbook
            string workbookUrl = "https://example.com/source.xlsx";

            // URL of the destination (cloud storage) where the workbook will be uploaded
            string uploadUrl = "https://cloudstorage.example.com/container/target.xlsx";

            // Download the workbook into a byte array using HttpClient
            byte[] workbookData;
            using (HttpClient httpClient = new HttpClient())
            {
                workbookData = httpClient.GetByteArrayAsync(workbookUrl).Result;
            }

            // Load the workbook from the downloaded data
            Workbook workbook;
            using (MemoryStream ms = new MemoryStream(workbookData))
            {
                workbook = new Workbook(ms);
            }

            // ------------------------------------------------------------
            // Apply a simple background style to a specific range (A1:B2)
            // ------------------------------------------------------------
            // Get the style of the first cell in the range
            Style style = workbook.Worksheets[0].Cells["A1"].GetStyle();

            // Set a solid background color (e.g., LightGray)
            style.Pattern = BackgroundType.Solid;
            style.ForegroundColor = System.Drawing.Color.LightGray;

            // Apply the style to the desired range
            Aspose.Cells.Range range = workbook.Worksheets[0].Cells.CreateRange("A1:B2");
            range.ApplyStyle(style, new StyleFlag { All = true });

            // Save the modified workbook into a memory stream
            byte[] modifiedData;
            using (MemoryStream outStream = new MemoryStream())
            {
                workbook.Save(outStream, SaveFormat.Xlsx);
                modifiedData = outStream.ToArray();
            }

            // Upload the modified workbook to cloud storage using HTTP PUT
            using (HttpClient client = new HttpClient())
            {
                using (ByteArrayContent content = new ByteArrayContent(modifiedData))
                {
                    content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                    HttpResponseMessage response = client.PutAsync(uploadUrl, content).Result;
                    response.EnsureSuccessStatusCode();
                }
            }

            Console.WriteLine("Workbook processed and uploaded successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
