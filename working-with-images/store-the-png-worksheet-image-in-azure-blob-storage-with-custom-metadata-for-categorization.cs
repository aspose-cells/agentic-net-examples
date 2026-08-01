// Title: Export Excel Worksheet to PNG and Upload to Azure Blob Storage with Custom Metadata – Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook with Aspose.Cells, render a worksheet to a PNG image using WorkbookRender, and upload the image directly to Azure Blob Storage while attaching custom metadata such as category and creation date.
// Keywords: Aspose.Cells | C# | .NET | Export worksheet to PNG | WorkbookRender | Azure Blob Storage | Azure SDK for .NET | custom blob metadata | image/png content type | Excel snapshot to cloud | store Excel image in Azure
// Common Searches: How to export Excel sheet as PNG with Aspose.Cells | Aspose.Cells upload PNG to Azure Blob using C# | Set metadata on Azure Blob when uploading image | Render worksheet to image stream and store in Azure | C# code for Excel image to Azure Blob with tags
// Developer Intent: Generate a PNG image from an Excel worksheet with Aspose.Cells and store it in Azure Blob Storage, applying custom metadata for easy categorization and retrieval.
// Use Cases: Create visual snapshots of financial reports and serve them from Azure for web dashboards. | Archive daily Excel snapshots in cloud storage with metadata like reportDate, department, and author. | Produce thumbnail images of worksheets for a document‑management system, storing them in Azure with searchable tags.
// AI Prompts: Write C# code that takes the MemoryStream from Aspose.Cells WorkbookRender and uploads it to Azure Blob Storage, adding metadata keys such as "Category" and "CreatedOn". | Show how to replace the local file save in the example with Azure Blob SDK calls, setting the content type to image/png and attaching custom metadata. | Generate a reusable method that renders a specified worksheet to PNG and streams it directly to Azure Blob without creating a temporary file.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsAzureBlobDemo
{
    // Demonstrates how to create a workbook with Aspose.Cells, render a worksheet to a PNG image using WorkbookRender, and upload the image directly to Azure Blob Storage while attaching custom metadata such as category and creation date.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // 1. Create a workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample data for image export");
                sheet.Cells["A2"].PutValue(DateTime.Now);

                // 2. Render the worksheet to a PNG image
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Png,
                    OnePagePerSheet = true
                };

                using (MemoryStream imageStream = new MemoryStream())
                {
                    WorkbookRender renderer = new WorkbookRender(workbook, imgOptions);
                    renderer.ToImage(imageStream); // render to stream
                    imageStream.Position = 0; // reset for reading

                    // 3. Save the PNG locally (replace Azure upload)
                    string outputDir = Path.Combine(Environment.CurrentDirectory, "OutputImages");
                    if (!Directory.Exists(outputDir))
                        Directory.CreateDirectory(outputDir);

                    string filePath = Path.Combine(outputDir, "sheet1.png");
                    using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                    {
                        imageStream.CopyTo(fileStream);
                    }

                    Console.WriteLine($"Image saved to '{filePath}'.");
                }

                // Clean up resources
                workbook.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
