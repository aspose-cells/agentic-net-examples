using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsRemoteLoadDemo
{
    public class RemoteWorkbookProcessor
    {
        // URL of the remote Excel file
        private const string RemoteExcelUrl = "https://example.com/sample.xlsx";

        // Optional local fallback file
        private const string LocalFallbackPath = "sample.xlsx";

        public static async Task RunAsync()
        {
            try
            {
                // Try to download the workbook from the remote URL
                using (HttpClient httpClient = new HttpClient())
                {
                    HttpResponseMessage response = await httpClient.GetAsync(RemoteExcelUrl);
                    if (response.IsSuccessStatusCode)
                    {
                        using (Stream responseStream = await response.Content.ReadAsStreamAsync())
                        {
                            ProcessWorkbook(new Workbook(responseStream));
                            return;
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Remote file not found (status {(int)response.StatusCode}).");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error downloading remote workbook: {ex.Message}");
            }

            // Fallback to a local file if it exists
            if (File.Exists(LocalFallbackPath))
            {
                try
                {
                    ProcessWorkbook(new Workbook(LocalFallbackPath));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading local workbook: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("No local workbook available for processing.");
            }
        }

        // Centralized workbook processing logic
        private static void ProcessWorkbook(Workbook workbook)
        {
            try
            {
                // Example processing: read the value of cell A1 from the first worksheet
                Worksheet firstSheet = workbook.Worksheets[0];
                string cellValue = firstSheet.Cells["A1"].StringValue;
                Console.WriteLine($"Value of A1: {cellValue}");

                // Optionally, save the workbook to a memory stream (e.g., to send elsewhere)
                using (MemoryStream outStream = new MemoryStream())
                {
                    workbook.Save(outStream, SaveFormat.Xlsx);
                    outStream.Position = 0;
                    Console.WriteLine($"Workbook saved to memory stream, length = {outStream.Length} bytes");
                }
            }
            finally
            {
                // Ensure resources are released
                workbook.Dispose();
            }
        }

        // Entry point for demonstration
        public static void Main()
        {
            RunAsync().GetAwaiter().GetResult();
        }
    }
}