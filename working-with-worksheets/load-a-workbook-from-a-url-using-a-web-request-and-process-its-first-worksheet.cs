using System;
using System.IO;
using System.Net.Http;
using Aspose.Cells;

namespace WorkbookProcessing
{
    class LoadWorkbookFromUrl
    {
        static void Main()
        {
            try
            {
                // URL of the Excel file to download
                string fileUrl = "https://example.com/sample.xlsx";

                // Download the file into a memory stream
                using (HttpClient client = new HttpClient())
                {
                    HttpResponseMessage response = client.GetAsync(fileUrl).Result;
                    response.EnsureSuccessStatusCode();

                    using (Stream httpStream = response.Content.ReadAsStreamAsync().Result)
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        httpStream.CopyTo(memoryStream);
                        memoryStream.Position = 0; // Reset stream position for reading

                        // Load the workbook from the memory stream
                        Workbook workbook = new Workbook(memoryStream);

                        // Access the first worksheet
                        Worksheet firstSheet = workbook.Worksheets[0];

                        // Read value of cell A1
                        Console.WriteLine("Value in A1: " + firstSheet.Cells["A1"].StringValue);

                        // Write a new value to cell B2
                        firstSheet.Cells["B2"].PutValue("Processed");

                        // Save the modified workbook
                        string outputPath = "ProcessedWorkbook.xlsx";
                        workbook.Save(outputPath, SaveFormat.Xlsx);
                        Console.WriteLine($"Workbook saved to {outputPath}");
                    }
                }
            }
            catch (HttpRequestException ex)
            {
                Console.WriteLine($"Error downloading file: {ex.Message}");
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine($"File not found: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}