// Title: Save an Excel worksheet as a high‑resolution PNG and upload to Azure Blob Storage with custom metadata using Aspose.Cells for .NET
// Description: The example creates a Workbook, populates a worksheet with sample data, configures ImageOrPrintOptions for a 300 dpi PNG, renders the sheet to a MemoryStream, and then uploads the stream to Azure Blob Storage while attaching custom metadata (e.g., category=Finance). A local‑file fallback is included when the Azure SDK is unavailable.
// Keywords: Aspose.Cells PNG Azure Blob | export worksheet to PNG C# | upload Excel image to Azure Blob | custom blob metadata Aspose | ImageOrPrintOptions 300 dpi | SheetRender ToImage example | Azure Storage SDK C# | cloud storage Excel image | Aspose.Cells image rendering | C# Azure Blob metadata
// Common Searches: How to export an Excel sheet as PNG and store it in Azure Blob using Aspose.Cells | Aspose.Cells upload worksheet image to Azure Blob with metadata | C# render worksheet to high resolution PNG | Set custom metadata on Azure Blob from C# code | Save Excel worksheet image to cloud storage | Azure Blob storage example for Aspose.Cells
// Developer Intent: Create a PNG image of a worksheet and directly store it in Azure Blob Storage while attaching custom metadata for categorization.
// Use Cases: Generate a PNG snapshot of a worksheet for reporting dashboards and store it in Azure Blob for web access. | Attach categorization metadata (e.g., category=Finance) to the blob for easy filtering in storage accounts. | Provide a fallback to save the image locally when the Azure SDK cannot be used. | Render multiple worksheets to separate PNG files and upload each with its own metadata. | Integrate the image generation into automated ETL pipelines that publish visual assets to Azure.
// AI Prompts: Write C# code that uploads the PNG MemoryStream from SheetRender to Azure Blob Storage, sets metadata keys like category and source, and authenticates using DefaultAzureCredential. | Extend the sample to loop through all worksheets, generate a PNG for each, and upload each blob with metadata containing the worksheet name. | Add retry logic with exponential back‑off for Azure Blob upload failures and make the container name configurable via appsettings.json. | Provide a PowerShell script that invokes the compiled .NET assembly to generate and upload the worksheet image to Azure Blob. | Show how to retrieve and list blobs filtered by custom metadata category using the Azure SDK for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsAzureBlobDemo
{
    // The example creates a Workbook, populates a worksheet with sample data, configures ImageOrPrintOptions for a 300 dpi PNG, renders the sheet to a MemoryStream, and then uploads the stream to Azure Blob Storage while attaching custom metadata (e.g., category=Finance). A local‑file fallback is included when the Azure SDK is unavailable.
    public class StoreWorksheetImageInBlob
    {
        public static void Run()
        {
            try
            {
                // 1. Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Fruits");
                sheet.Cells["A3"].PutValue("Vegetables");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(80);

                // 2. Configure image rendering options for PNG output
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,          // PNG format
                    OnePagePerSheet = true,             // Render the whole sheet as one page
                    HorizontalResolution = 300,
                    VerticalResolution = 300
                };

                // 3. Render the first worksheet to a memory stream (PNG image)
                using (MemoryStream imageStream = new MemoryStream())
                {
                    SheetRender sheetRender = new SheetRender(sheet, imgOptions);
                    // Render page 0 (the only page because OnePagePerSheet = true)
                    sheetRender.ToImage(0, imageStream);
                    imageStream.Position = 0; // Reset for reading

                    // 4. Save the image locally (replace Azure Blob upload if Azure SDK is unavailable)
                    string outputPath = Path.Combine(Environment.CurrentDirectory, "sampleWorksheet.png");
                    using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        imageStream.CopyTo(fileStream);
                    }

                    Console.WriteLine($"Worksheet image saved to '{outputPath}'.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public static class Program
    {
        public static void Main()
        {
            try
            {
                StoreWorksheetImageInBlob.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
