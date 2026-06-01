using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace ImageMarkerDemo
{
    class ImageMarkerTemplateDemo
    {
        static void Main()
        {
            try
            {
                // Sample base64 image string (a small red dot PNG)
                string base64Image = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";

                // Convert base64 string to byte array
                byte[] imageBytes = Convert.FromBase64String(base64Image);

                // Create a DataTable that will serve as the data source for the image marker
                DataTable photoTable = new DataTable("Photos");
                photoTable.Columns.Add("Photo", typeof(byte[]));
                photoTable.Rows.Add(imageBytes);

                // Create a workbook that will act as the template
                Workbook templateWorkbook = new Workbook();
                Worksheet sheet = templateWorkbook.Worksheets[0];

                // Insert an image smart marker into cell A1
                // The marker syntax "&=Photo" tells Aspose.Cells to replace it with the image from the "Photo" column
                sheet.Cells["A1"].PutValue("&=Photo");

                // Initialize WorkbookDesigner with the template workbook
                WorkbookDesigner designer = new WorkbookDesigner(templateWorkbook);

                // Set the data source (the DataTable) for the smart markers
                designer.SetDataSource(photoTable);

                // Process the smart markers – this will embed the image into the worksheet
                designer.Process();

                // Save the result as HTML with images embedded as Base64 strings
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportImagesAsBase64 = true   // Embed images directly in the HTML using Base64
                };

                string outputPath = "ImageMarkerOutput.html";

                // Save the processed workbook
                designer.Workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"Workbook processed and saved as HTML with embedded Base64 images: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}