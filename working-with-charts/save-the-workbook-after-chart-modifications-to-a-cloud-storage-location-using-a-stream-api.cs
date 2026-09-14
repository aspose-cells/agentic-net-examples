// Title: Asynchronously save an Aspose.Cells workbook with an updated chart to a MemoryStream and write it to a file or cloud storage in C#
// AI Prompts: Generate C# code that creates a workbook, populates it with data, adds a column chart, changes the chart title, saves the workbook to a MemoryStream in XLSX format, and then uses async I/O to copy the stream to a given file path. | Show how to modify a chart in an Aspose.Cells workbook and persist the changes by streaming the workbook to a cloud storage service (e.g., Azure Blob, AWS S3) with an asynchronous upload method in .NET. | Provide an async helper method that receives a Stream and a destination URI or file path, ensures the target container or directory exists, and uploads the stream using the appropriate SDK.
// Common Searches: c# Aspose.Cells export modified chart to MemoryStream and upload to Azure Blob storage | how to save Aspose.Cells workbook as XLSX using async stream to Google Cloud Storage | example of updating chart title in Aspose.Cells and writing workbook to a file with async I/O | Aspose.Cells save workbook to stream then copy to Amazon S3 using .NET async
// Tags: Aspose.Cells save workbook to MemoryStream | chart title update Aspose.Cells C# | async stream write to file .NET | upload Excel stream to Azure Blob using Aspose.Cells | export XLSX via stream API cloud storage

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Charts; // Added for Chart and ChartType

namespace AsposeCellsCloudSaveDemo
{
    // The sample creates a new workbook, adds sample data, inserts a column chart, updates the chart title, saves the workbook to a MemoryStream in XLSX format, resets the stream position, and then asynchronously writes the stream to a local file while ensuring the target directory exists—ready to be swapped for any cloud storage upload implementation.
    class Program
    {
        // Local file path to save the workbook (replace with desired location)
        private const string OutputFilePath = "ModifiedWorkbook.xlsx";

        static async Task Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and add sample data
                using (Workbook workbook = new Workbook())
                {
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Cells["A1"].PutValue("Category");
                    sheet.Cells["A2"].PutValue("Apple");
                    sheet.Cells["A3"].PutValue("Banana");
                    sheet.Cells["A4"].PutValue("Cherry");
                    sheet.Cells["B1"].PutValue("Value");
                    sheet.Cells["B2"].PutValue(30);
                    sheet.Cells["B3"].PutValue(45);
                    sheet.Cells["B4"].PutValue(25);

                    // 2. Add a column chart
                    int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                    Chart chart = sheet.Charts[chartIndex];
                    chart.NSeries.Add("B2:B4", true);
                    chart.NSeries.CategoryData = "A2:A4";
                    chart.Title.Text = "Fruit Sales";

                    // 3. Modify the chart (example: change title)
                    chart.Title.Text = "Updated Fruit Sales";

                    // 4. Save the workbook to a memory stream
                    using (MemoryStream workbookStream = new MemoryStream())
                    {
                        workbook.Save(workbookStream, SaveFormat.Xlsx);
                        workbookStream.Position = 0; // Reset stream position before saving to file

                        // 5. Save the stream to a local file
                        await SaveStreamToFileAsync(workbookStream, OutputFilePath);
                    }
                }

                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(OutputFilePath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Saves a stream to a file, creating or overwriting the target file.
        private static async Task SaveStreamToFileAsync(Stream dataStream, string filePath)
        {
            try
            {
                // Ensure the directory exists
                string? directory = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Write the stream to the file
                using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    await dataStream.CopyToAsync(fileStream);
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to save file '{filePath}': {ex.Message}");
                throw;
            }
        }
    }
}
