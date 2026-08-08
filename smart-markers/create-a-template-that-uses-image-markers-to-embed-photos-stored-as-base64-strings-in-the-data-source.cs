// Title: Embed Base64 Images with Smart Markers in Aspose.Cells .NET – JSON to HTML
// Description: Demonstrates how to create an Excel template containing an image smart marker, supply a Base64‑encoded picture through a JSON data source, process the marker with WorkbookDesigner, and save the result as HTML with the image embedded as Base64. No external image files are required at runtime.
// Keywords: Aspose.Cells image smart marker | Base64 image embedding C# | JSON data source Aspose.Cells | Export Excel to HTML with inline images | Aspose.Cells WorkbookDesigner example | C# generate HTML from Excel | smart marker picture replacement
// Common Searches: Aspose.Cells replace smart marker with Base64 image | C# embed image in Excel template using JSON | Export Aspose.Cells workbook to HTML with inline Base64 pictures | How to use image smart markers in Aspose.Cells .NET | Load Base64 photo into Excel cell via smart marker
// Developer Intent: Create an HTML report where a smart marker is automatically replaced by a Base64‑encoded image supplied via a JSON data source using Aspose.Cells for .NET.
// Use Cases: Generate personalized PDFs or web pages that include user avatars stored as Base64 strings in a database. | Automate conversion of Excel templates with image placeholders into web‑ready HTML without managing separate image files. | Build data‑driven dashboards where charts and photos are injected into spreadsheets from API responses.
// AI Prompts: Write C# code that uses Aspose.Cells WorkbookDesigner to bind a JSON object containing a Base64 image to an image smart marker and export the workbook as HTML with embedded images. | Explain the required HtmlSaveOptions settings to output Base64‑encoded images when saving an Aspose.Cells workbook to HTML. | Provide step‑by‑step instructions for converting a local JPEG to a Base64 string, embedding it in JSON, and linking it to an "&=$Photo" smart marker in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsImageMarkerDemo
{
    // Demonstrates how to create an Excel template containing an image smart marker, supply a Base64‑encoded picture through a JSON data source, process the marker with WorkbookDesigner, and save the result as HTML with the image embedded as Base64. No external image files are required at runtime.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook that will serve as the template
                Workbook templateWorkbook = new Workbook();
                Worksheet sheet = templateWorkbook.Worksheets[0];

                // Place an image smart marker in cell A1.
                // The marker syntax "&=$Photo" tells Aspose.Cells to replace it with an image from the data source.
                sheet.Cells["A1"].PutValue("&=$Photo");

                // Load an image file and convert it to a Base64 string.
                // In a real scenario the Base64 string could come from a database or an API.
                string imagePath = "photo.jpg";
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                    return;
                }

                byte[] imageBytes = File.ReadAllBytes(imagePath);
                string base64Image = Convert.ToBase64String(imageBytes);

                // Prepare JSON data source containing the Base64 image string.
                // The property name "Photo" must match the smart marker name.
                string jsonData = $"{{\"Photo\":\"{base64Image}\"}}";

                // Initialize WorkbookDesigner with the template workbook.
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = templateWorkbook
                };

                // Set the JSON data source.
                designer.SetJsonDataSource("Data", jsonData);

                // Process the smart markers; the image marker will be replaced with the image.
                designer.Process();

                // Save the result as HTML with images embedded as Base64.
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportImagesAsBase64 = true // Embed images directly in the HTML.
                };

                // The workbook now contains the image; save it.
                designer.Workbook.Save("Result.html", saveOptions);

                Console.WriteLine("HTML file with embedded Base64 image generated successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
