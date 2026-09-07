// Title: Render an Excel worksheet to PNG and store it in Azure Blob Storage with custom metadata using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, renders a specified worksheet page to a PNG stream, and uploads the stream to an Azure Blob container while attaching custom metadata key‑value pairs. | Modify the Aspose.Cells worksheet‑to‑image example to bypass the local file system and directly write the PNG image to Azure Blob Storage, setting metadata such as Category, Source, and CreatedBy on the blob. | Create a reusable C# method that accepts a workbook path, worksheet index, Azure Blob connection string, container name, and a dictionary of metadata, then returns the URL of the uploaded PNG image.
// Common Searches: Aspose.Cells C# render worksheet to PNG and upload to Azure Blob with metadata | How to add custom metadata to a blob when saving an Aspose.Cells rendered image | Store Excel sheet image in Azure Blob Storage using Aspose.Cells .NET SDK | C# example for streaming Aspose.Cells PNG output directly to Azure Blob | Set metadata on Azure Blob for images generated from Excel worksheets
// Tags: Aspose.Cells PNG image generation | Azure Blob upload PNG with metadata | C# stream image to Azure Blob storage | custom metadata on Azure blob for generated images | Excel to image conversion Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsAzureBlobExample
{
    // The program loads 'Sample.xlsx' with Aspose.Cells, renders the first worksheet page to a PNG image using SheetRender, writes the image to a MemoryStream, then copies the stream to a local file 'WorksheetImage.png', handling missing file and runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string excelFilePath = "Sample.xlsx";

            // Verify that the Excel file exists to avoid FileNotFoundException
            if (!File.Exists(excelFilePath))
            {
                Console.WriteLine($"Error: The file '{excelFilePath}' was not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(excelFilePath);

                // Choose the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Set image rendering options (default format is PNG)
                ImageOrPrintOptions options = new ImageOrPrintOptions();

                // Render the worksheet to a PNG image in memory using SheetRender
                using (MemoryStream imageStream = new MemoryStream())
                {
                    SheetRender render = new SheetRender(worksheet, options);

                    // Render the first page (page index starts at 0)
                    render.ToImage(0, imageStream);
                    imageStream.Position = 0; // Reset stream position for further use

                    // Save the image locally
                    string outputImagePath = "WorksheetImage.png";
                    using (FileStream fileStream = new FileStream(outputImagePath, FileMode.Create, FileAccess.Write))
                    {
                        imageStream.CopyTo(fileStream);
                    }

                    Console.WriteLine($"Worksheet image saved to '{outputImagePath}'.");
                }
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display the error
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
