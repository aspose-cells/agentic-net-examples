// Title: Download an Excel workbook from a URL, replace placeholders in all TextBox shapes, and upload the modified file to cloud storage using Aspose.Cells for .NET
// Common Searches: Aspose.Cells replace {{Name}} placeholder in TextBox shapes after loading workbook from HTTP stream | C# download Excel file from URL, edit TextBox text, and upload back to cloud storage | How to iterate over worksheet shapes and modify TextBox content with Aspose.Cells in .NET | Upload modified Excel workbook to REST endpoint using HttpClient and Aspose.Cells | Process Excel workbook in memory stream to replace tags in shapes with Aspose.Cells
// Tags: replace text in Excel TextBox shapes Aspose.Cells | download and load Excel workbook via HttpClient Aspose.Cells | upload Excel file via HTTP PUT C# | iterate worksheet shapes Aspose.Cells | memory stream Excel processing Aspose.Cells

using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example downloads an Excel workbook from a remote URL into a memory stream, uses Aspose.Cells to iterate through each worksheet and its shapes, replaces {{Name}} and {{Date}} placeholders in TextBox shapes, saves the updated workbook back to a memory stream, and uploads the result to a cloud storage endpoint with an HTTP PUT request.
class Program
{
    // Entry point
    static async Task Main(string[] args)
    {
        // URLs for source workbook and destination upload
        string sourceUrl = "https://example.com/source.xlsx";
        string destinationUrl = "https://cloudstorage.example.com/upload/modified.xlsx";

        try
        {
            // Download the workbook into a memory stream
            using (HttpClient httpClient = new HttpClient())
            using (Stream downloadStream = await httpClient.GetStreamAsync(sourceUrl))
            using (MemoryStream workbookStream = new MemoryStream())
            {
                await downloadStream.CopyToAsync(workbookStream);
                workbookStream.Position = 0; // Reset stream position for reading

                // Load workbook from the memory stream
                Workbook workbook = new Workbook(workbookStream);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all shapes in the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        try
                        {
                            // Process shapes that contain text (e.g., TextBox)
                            if (!string.IsNullOrEmpty(shape.Text))
                            {
                                // Get current text
                                string originalText = shape.Text;

                                // Example tag replacement
                                string replacedText = originalText
                                    .Replace("{{Name}}", "John Doe")
                                    .Replace("{{Date}}", DateTime.Now.ToShortDateString());

                                // Update the TextBox text if any changes were made
                                if (!originalText.Equals(replacedText))
                                {
                                    shape.Text = replacedText;
                                }
                            }
                        }
                        catch (Exception shapeEx)
                        {
                            Console.Error.WriteLine($"Shape processing error in sheet '{sheet.Name}': {shapeEx.Message}");
                        }
                    }
                }

                // Save the modified workbook into a new memory stream
                using (MemoryStream outputStream = new MemoryStream())
                {
                    workbook.Save(outputStream, SaveFormat.Xlsx);
                    outputStream.Position = 0; // Reset for reading

                    // Upload the modified workbook to cloud storage via HTTP PUT
                    using (HttpContent content = new StreamContent(outputStream))
                    {
                        content.Headers.ContentType = new MediaTypeHeaderValue(
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

                        HttpResponseMessage response = await httpClient.PutAsync(destinationUrl, content);
                        response.EnsureSuccessStatusCode();
                    }
                }
            }

            Console.WriteLine("Workbook processed and uploaded successfully.");
        }
        catch (HttpRequestException httpEx)
        {
            Console.Error.WriteLine($"HTTP error: {httpEx.Message}");
        }
        catch (IOException ioEx)
        {
            Console.Error.WriteLine($"I/O error: {ioEx.Message}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
