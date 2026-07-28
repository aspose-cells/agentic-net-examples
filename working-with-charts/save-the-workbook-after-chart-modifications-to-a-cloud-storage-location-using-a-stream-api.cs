// Title: Save an Aspose.Cells Workbook with a Chart to Cloud Storage via MemoryStream (C#)
// Description: Demonstrates creating a workbook, adding a column chart, saving it to a MemoryStream in XLSX format, and shows where to plug in Azure Blob, Amazon S3, or Google Cloud Storage SDKs to upload the stream directly to the cloud.
// Keywords: Aspose.Cells | C# | MemoryStream | chart | save workbook | cloud upload | Azure Blob Storage | Amazon S3 | Google Cloud Storage | XLSX | Excel file upload
// Common Searches: Aspose.Cells upload workbook to Azure Blob using MemoryStream | Save Excel file with chart to Amazon S3 in C# | Stream Aspose.Cells workbook to Google Cloud Storage | C# example for saving Aspose.Cells chart to cloud storage | How to use MemoryStream with Aspose.Cells for cloud upload
// Developer Intent: Upload a chart‑enhanced Excel workbook directly to a cloud storage service using a stream.
// Use Cases: Pass the MemoryStream to Azure Blob SDK (BlobClient) for immediate storage in a container. | Send the stream to AWS S3 via the AmazonS3Client PutObjectAsync method. | Write the stream to a Google Cloud Storage bucket using Google.Cloud.Storage.V1 StorageClient.
// AI Prompts: Generate C# code that replaces the local file write with an Azure Blob Storage upload using workbookStream. | Show how to upload the MemoryStream returned by workbook.Save to an Amazon S3 bucket with the AWS SDK for .NET. | Create a method that accepts a MemoryStream and stores it in a Google Cloud Storage bucket using the Google.Cloud.Storage.V1 library.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCloudSaveDemo
{
    // Demonstrates creating a workbook, adding a column chart, saving it to a MemoryStream in XLSX format, and shows where to plug in Azure Blob, Amazon S3, or Google Cloud Storage SDKs to upload the stream directly to the cloud.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook
                Workbook workbook = new Workbook();

                // 2. Populate sample data
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["A4"].PutValue("Cherry");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["B3"].PutValue(45);
                sheet.Cells["B4"].PutValue(25);

                // 3. Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // 4. Set chart title
                chart.Title.Text = "Fruit Sales";

                // 5. Save the workbook to a memory stream
                using (MemoryStream workbookStream = new MemoryStream())
                {
                    workbook.Save(workbookStream, SaveFormat.Xlsx);
                    workbookStream.Position = 0; // Reset for reading

                    // 6. Save the stream to a local file (replace with Azure upload if SDK is available)
                    string outputPath = Path.Combine(Environment.CurrentDirectory, "charts_demo.xlsx");

                    // Ensure the directory exists
                    string outputDir = Path.GetDirectoryName(outputPath);
                    if (!Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Write the stream to the file
                    using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        workbookStream.CopyTo(fileStream);
                    }

                    Console.WriteLine($"Workbook saved locally at '{outputPath}'.");
                }

                // Clean up
                workbook.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
