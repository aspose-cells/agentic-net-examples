using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ImageMarkerTemplate
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Place an image smart marker in cell A1.
        // The marker will be replaced by the image data from the JSON data source.
        sheet.Cells["A1"].PutValue("&=$Data.Photo");

        // Sample PNG image encoded as Base64 (a tiny red dot)
        string base64Image = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAusB9W5XK9cAAAAASUVORK5CYII=";

        // Build JSON where the Photo field contains the Base64 string
        string jsonData = $"{{\"Photo\":\"{base64Image}\"}}";

        // Initialize WorkbookDesigner with the workbook template
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Set the JSON data source; the name "Data" matches the marker prefix
        designer.SetJsonDataSource("Data", jsonData);

        // Process the smart markers – the image marker will be replaced with the image
        designer.Process();

        // Save the result as HTML with images embedded as Base64 strings
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportImagesAsBase64 = true   // Embed images directly in the HTML
        };
        workbook.Save("ImageMarkerResult.html", saveOptions);
    }
}