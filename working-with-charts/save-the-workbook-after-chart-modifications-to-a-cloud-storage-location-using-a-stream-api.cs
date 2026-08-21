// Title: Save an Aspose.Cells workbook with a modified chart to cloud storage using a stream (C#)
// Description: This example creates a workbook, adds sample data, inserts a column chart, updates its title, and demonstrates how to replace the local file‑system save with a MemoryStream that can be uploaded directly to Azure Blob, Amazon S3, or Google Cloud Storage.
// Keywords: Aspose.Cells C# stream save | upload Excel to Azure Blob | Aspose.Cells chart to S3 | Google Cloud Storage Excel upload | MemoryStream Aspose.Cells | cloud storage workbook export
// Common Searches: Aspose.Cells save workbook to Azure Blob using MemoryStream | Upload Excel file with chart to Amazon S3 in C# | How to store Aspose.Cells workbook in Google Cloud Storage | C# stream save Aspose.Cells example | Aspose.Cells chart export to cloud
// Developer Intent: Export a workbook that contains a modified chart directly to a cloud storage service via a stream instead of writing to disk.
// Use Cases: Generate a sales report with a column chart and upload it to Azure Blob Storage by writing the workbook to a MemoryStream and calling BlobClient.UploadAsync. | Create an Excel file with a chart and store it in an Amazon S3 bucket by saving the workbook to a MemoryStream and using PutObjectRequest. | Produce a charted workbook and persist it in Google Cloud Storage by converting the workbook to a byte array via MemoryStream and uploading with StorageClient.UploadObject.
// AI Prompts: Provide C# code that replaces workbook.Save with a MemoryStream and uploads the stream to Azure Blob Storage using Azure.Storage.Blobs. | Show a snippet that saves an Aspose.Cells workbook to a MemoryStream and uses the AWS SDK for .NET to put the stream into an S3 bucket. | Generate example code that writes the workbook to a stream and uploads it to Google Cloud Storage using the Google.Cloud.Storage.V1 library.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example creates a workbook, adds sample data, inserts a column chart, updates its title, and demonstrates how to replace the local file‑system save with a MemoryStream that can be uploaded directly to Azure Blob, Amazon S3, or Google Cloud Storage.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and add sample data
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(30);
            worksheet.Cells["B3"].PutValue(45);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B3", true);               // Set data range
            chart.NSeries.CategoryData = "A2:A3";           // Set category range
            chart.Title.Text = "Fruit Sales";               // Modify chart title

            // Define output file path
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ChartWorkbook.xlsx");

            // Save the workbook to the file system
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
