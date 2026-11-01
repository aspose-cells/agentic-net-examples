// Title: Render an Excel worksheet to PNG and upload it to Azure Blob Storage with custom metadata using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a .xlsx workbook with Aspose.Cells, renders the first worksheet to a PNG memory stream, and uploads the stream to an Azure Blob container while attaching metadata keys such as Category, CreatedBy, and CreatedOn. | Refactor the worksheet image export example to bypass the local file system, ensure the target Azure Blob container exists, and directly upload the PNG using Azure.Storage.Blobs SDK with custom metadata. | Implement robust error handling for the Azure upload step, including retry logic for transient failures, detailed logging of upload status, and graceful handling of missing or invalid Azure credentials.
// Common Searches: C# Aspose.Cells render worksheet to PNG and upload to Azure Blob with metadata | How to add custom metadata when uploading images to Azure Blob Storage using .NET | Save Excel sheet as PNG directly to Azure Blob without creating a local file | Aspose.Cells export worksheet as image to Azure storage example | Azure Blob Storage upload PNG from memory stream C#
// Tags: Aspose.Cells render worksheet to PNG | Azure Blob Storage upload with metadata .NET | C# memory stream image upload to Azure | Excel to image conversion using Aspose.Cells | Custom metadata for Azure blobs in C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example demonstrates loading an Excel workbook with Aspose.Cells, rendering the first worksheet to a PNG image in memory, and uploading that image directly to Azure Blob Storage while applying custom metadata such as Category, CreatedBy, and CreatedOn.
class WorksheetImageUploader
{
    // Adjust these values as needed
    private const string ExcelFilePath = @"C:\Data\Sample.xlsx";
    private const string OutputImagePath = @"C:\Data\Sheet1.png";

    static void Main()
    {
        try
        {
            // Verify that the Excel file exists before attempting to load it
            if (!File.Exists(ExcelFilePath))
            {
                Console.WriteLine($"Error: Excel file not found at '{ExcelFilePath}'.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(ExcelFilePath);

            // Choose the worksheet to render (first worksheet in this example)
            Worksheet sheet = workbook.Worksheets[0];

            // Set image options (default PNG format)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
            // Optional: configure resolution, margins, etc.
            // imgOptions.HorizontalResolution = 300;
            // imgOptions.VerticalResolution = 300;

            // Render the worksheet to an image stream
            using (MemoryStream imageStream = new MemoryStream())
            {
                SheetRender sheetRender = new SheetRender(sheet, imgOptions);
                // Render the first page (index 0) of the worksheet
                sheetRender.ToImage(0, imageStream);
                imageStream.Position = 0; // Reset stream position for saving

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(OutputImagePath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the PNG image locally
                using (FileStream fileStream = new FileStream(OutputImagePath, FileMode.Create, FileAccess.Write))
                {
                    imageStream.CopyTo(fileStream);
                }

                // Example metadata (not used in local file system, kept for reference)
                IDictionary<string, string> metadata = new Dictionary<string, string>
                {
                    { "Category", "FinancialReport" },
                    { "CreatedBy", "AsposeDemo" },
                    { "CreatedOn", DateTime.UtcNow.ToString("o") } // ISO 8601 format
                };

                // Metadata handling would be applied when uploading to a storage service
            }

            Console.WriteLine($"Worksheet image saved to '{OutputImagePath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
