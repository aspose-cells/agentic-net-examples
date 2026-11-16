// Title: Download an Excel file via HttpWebRequest, load it with Aspose.Cells, edit the first worksheet, and save the changes in C#
// Common Searches: C# download Excel file from URL and open with Aspose.Cells Workbook constructor | Read and write cells in the first worksheet after loading workbook from a web response stream using Aspose.Cells | Aspose.Cells fallback to local file when web request fails in .NET | Save modified workbook to a new file after editing cells with Aspose.Cells in C# | Handle WebException while loading Excel workbook with Aspose.Cells
// Tags: HttpWebRequest Excel download Aspose.Cells | Workbook initialization from response stream | first worksheet cell A1 read Aspose.Cells | cell B2 write Aspose.Cells | save workbook as new file Aspose.Cells | exception handling fallback local workbook Aspose.Cells

using System;
using System.IO;
using System.Net;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example demonstrates how to use HttpWebRequest to fetch an Excel file from a remote URL, load the workbook directly from the response stream with Aspose.Cells, and gracefully fall back to a local file if the download fails. It accesses the first worksheet, reads the value of cell A1, writes "Processed" to cell B2, and saves the updated workbook as a new file, while handling relevant exceptions.
    class Program
    {
        static void Main()
        {
            Workbook workbook = null;
            string fileUrl = "https://example.com/sample.xlsx";
            string localPath = "sample.xlsx";

            // Try to download the workbook from the URL
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(fileUrl);
                request.Method = "GET";

                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                using (Stream responseStream = response.GetResponseStream())
                {
                    workbook = new Workbook(responseStream);
                }
            }
            catch (WebException ex)
            {
                Console.WriteLine($"Failed to download file from URL: {ex.Message}");

                // Fallback to a local file if it exists
                if (File.Exists(localPath))
                {
                    try
                    {
                        workbook = new Workbook(localPath);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine($"Failed to load local workbook: {e.Message}");
                        return;
                    }
                }
                else
                {
                    Console.WriteLine("No local workbook found. Exiting.");
                    return;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
                return;
            }

            // Process the workbook
            try
            {
                Worksheet sheet = workbook.Worksheets[0];

                // Read value of cell A1
                Cell cellA1 = sheet.Cells["A1"];
                Console.WriteLine("Value of A1: " + cellA1.StringValue);

                // Write a new value to cell B2
                sheet.Cells["B2"].PutValue("Processed");

                // Save the modified workbook
                string outputPath = "ProcessedSample.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
