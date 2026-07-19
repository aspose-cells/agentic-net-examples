// Title: Save an Aspose.Cells workbook with a modified chart to cloud storage using a MemoryStream (C#)
// Description: Creates a workbook, adds a column chart, updates the chart title, writes the workbook to a MemoryStream in XLSX format, resets the stream, and demonstrates how to pass the stream to a cloud‑storage upload method (e.g., Azure Blob, AWS S3, Google Cloud Storage).
// Keywords: Aspose.Cells C# save to MemoryStream | export chart workbook to XLSX stream | upload Aspose.Cells file to Azure Blob | stream workbook to Amazon S3 | Google Cloud Storage Aspose.Cells example | chart title change Aspose.Cells | .NET cloud storage stream API
// Common Searches: How to save an Aspose.Cells workbook to a MemoryStream | Aspose.Cells chart export to Azure Blob Storage | Upload XLSX stream to AWS S3 using C# | Google Cloud Storage upload for Aspose.Cells workbook | C# Aspose.Cells save workbook to cloud via stream
// Developer Intent: Generate a workbook with a customized chart, serialize it to a MemoryStream, and upload the stream to a cloud storage service.
// Use Cases: Automated sales reporting: create a chart, stream the XLSX, and store it in Azure Blob for downstream analytics. | Daily financial dashboards: update multiple charts, push the workbook to an Amazon S3 bucket for archival. | Cross‑region data sharing: generate a chart‑rich workbook and upload it to Google Cloud Storage for global access.
// AI Prompts: Provide Azure Blob Storage SDK code to upload the MemoryStream containing the workbook. | Show how to use the AWS SDK for .NET to stream the workbook to an S3 bucket. | Give a sample that uploads the workbook stream to Google Cloud Storage with the .NET client library.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts; // Required for Chart and ChartType

namespace AsposeCellsChartCloudSaveDemo
{
    // Creates a workbook, adds a column chart, updates the chart title, writes the workbook to a MemoryStream in XLSX format, resets the stream, and demonstrates how to pass the stream to a cloud‑storage upload method (e.g., Azure Blob, AWS S3, Google Cloud Storage).
    class Program
    {
        static void Main()
        {
            try
            {
                // Initialize a new workbook
                using (Workbook workbook = new Workbook())
                {
                    // Access the first worksheet
                    Worksheet worksheet = workbook.Worksheets[0];

                    // Populate sample data for the chart
                    worksheet.Cells["A1"].PutValue("Category");
                    worksheet.Cells["A2"].PutValue("Apple");
                    worksheet.Cells["A3"].PutValue("Banana");
                    worksheet.Cells["A4"].PutValue("Cherry");

                    worksheet.Cells["B1"].PutValue("Value");
                    worksheet.Cells["B2"].PutValue(30);
                    worksheet.Cells["B3"].PutValue(45);
                    worksheet.Cells["B4"].PutValue(25);

                    // Add a column chart to the worksheet
                    int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                    Chart chart = worksheet.Charts[chartIndex];

                    // Set the data source for the chart
                    chart.NSeries.Add("B2:B4", true);
                    chart.NSeries.CategoryData = "A2:A4";

                    // Change the chart title
                    chart.Title.Text = "Fruit Sales";

                    // Save the workbook to a memory stream in XLSX format
                    using (MemoryStream workbookStream = new MemoryStream())
                    {
                        workbook.Save(workbookStream, SaveFormat.Xlsx);
                        workbookStream.Position = 0; // Reset stream position

                        // Upload the stream to cloud storage (placeholder)
                        UploadToCloudStorage(workbookStream, "sample-workbook.xlsx");
                    }
                }
            }
            catch (Exception ex)
            {
                // Log or handle exceptions as needed
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Placeholder method representing cloud storage upload via a stream API.
        // Replace the body with actual SDK calls (e.g., Azure Blob, AWS S3, Google Cloud Storage).
        static void UploadToCloudStorage(Stream dataStream, string fileName)
        {
            // Example: using Azure Blob Storage SDK (commented out)
            // var blobServiceClient = new BlobServiceClient(connectionString);
            // var containerClient = blobServiceClient.GetBlobContainerClient(containerName);
            // var blobClient = containerClient.GetBlobClient(fileName);
            // blobClient.Upload(dataStream, overwrite: true);

            // For demonstration, simply indicate that the upload would occur here.
            Console.WriteLine($"[UploadToCloudStorage] Stream for '{fileName}' would be uploaded to cloud storage here.");
        }
    }
}
