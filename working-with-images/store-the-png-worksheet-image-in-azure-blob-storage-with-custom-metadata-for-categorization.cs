using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("Sample Data");

            // Set image rendering options for PNG output
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png
            };

            // Render the first worksheet to a PNG image in a memory stream
            WorkbookRender renderer = new WorkbookRender(workbook, imgOptions);
            using (MemoryStream imageStream = new MemoryStream())
            {
                renderer.ToImage(imageStream);          // Render to stream
                imageStream.Position = 0;               // Reset for reading

                // OPTIONAL: Save the image locally (replace Azure upload)
                string imagePath = "sampleWorksheet.png";
                using (FileStream file = new FileStream(imagePath, FileMode.Create, FileAccess.Write))
                {
                    imageStream.CopyTo(file);
                }

                // If Azure Blob Storage is required, ensure the Azure.Storage.Blobs package is referenced
                // and uncomment the following block after adding the appropriate using directive.
                /*
                string connectionString = "<YOUR_AZURE_BLOB_CONNECTION_STRING>";
                string containerName   = "worksheet-images";
                string blobName        = "sampleWorksheet.png";

                var container = new Azure.Storage.Blobs.BlobContainerClient(connectionString, containerName);
                container.CreateIfNotExists();

                var blob = container.GetBlobClient(blobName);
                imageStream.Position = 0; // Reset before upload
                blob.Upload(imageStream, overwrite: true);

                var metadata = new Dictionary<string, string>
                {
                    { "category", "financial" },
                    { "createdby", "aspnet" }
                };
                blob.SetMetadata(metadata);
                */
            }

            // Save the workbook to disk (optional)
            string workbookPath = "output.xlsx";
            workbook.Save(workbookPath);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}