// Title: Load Excel from URL with Aspose.Cells, fallback to local file, edit first sheet (C#)
// Description: C# example that downloads an XLSX file via HttpClient, loads it into an Aspose.Cells Workbook, reads cell A1, writes the current timestamp to B2, and saves the modified workbook. If the download fails, a local copy is used as a fallback.
// Keywords: Aspose.Cells load workbook from URL | C# HttpClient Excel download | fallback to local Excel file | read cell A1 Aspose.Cells | write timestamp to cell B2 | save modified workbook C#
// Common Searches: Aspose.Cells load workbook from web address | C# download Excel stream with HttpClient and edit | use local Excel file when web download fails Aspose | write current date to cell using Aspose.Cells | save workbook after modification .NET
// Developer Intent: Download an Excel file, modify the first worksheet, and persist the changes, with a local fallback if the download fails.
// Use Cases: Automate daily report ingestion from a web service, add processing timestamp, and archive locally. | Retrieve a remote template, populate header data in the first sheet, and generate a customized workbook. | Fetch a spreadsheet from a cloud endpoint; if unavailable, use a cached copy for further data manipulation.
// AI Prompts: Generate C# code that uses Aspose.Cells to load an Excel workbook from a URL with error handling and a local fallback. | Show how to read cell A1, write DateTime.Now to cell B2, and save the workbook as a new file using Aspose.Cells. | Refactor the example to use async/await for the HttpClient download and workbook processing.

using System;
using System.IO;
using System.Net.Http;
using Aspose.Cells;

namespace AsposeCellsWebLoadDemo
{
    // C# example that downloads an XLSX file via HttpClient, loads it into an Aspose.Cells Workbook, reads cell A1, writes the current timestamp to B2, and saves the modified workbook. If the download fails, a local copy is used as a fallback.
    class Program
    {
        static void Main()
        {
            // URL of the Excel file to download
            const string fileUrl = "https://example.com/sample.xlsx";

            Workbook workbook = null;

            // Attempt to download the workbook from the web
            try
            {
                using (HttpClient httpClient = new HttpClient())
                using (Stream excelStream = httpClient.GetStreamAsync(fileUrl).Result)
                {
                    // Load the workbook from the downloaded stream
                    workbook = new Workbook(excelStream);
                }
            }
            catch (HttpRequestException httpEx)
            {
                Console.WriteLine($"Web download failed: {httpEx.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error during download: {ex.Message}");
            }

            // If download failed, try to load a local template file
            if (workbook == null)
            {
                const string localPath = "sample.xlsx";
                if (File.Exists(localPath))
                {
                    try
                    {
                        workbook = new Workbook(localPath);
                        Console.WriteLine("Loaded workbook from local file.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to load local workbook: {ex.Message}");
                        return;
                    }
                }
                else
                {
                    Console.WriteLine("Neither web download nor local file is available. Exiting.");
                    return;
                }
            }

            // Process the workbook
            try
            {
                Worksheet firstSheet = workbook.Worksheets[0];

                // Read the value of cell A1
                Console.WriteLine("Value in A1: " + firstSheet.Cells["A1"].StringValue);

                // Write a timestamp into cell B2
                firstSheet.Cells["B2"].PutValue(DateTime.Now);

                // Save the modified workbook
                const string outputPath = "ProcessedWorkbook.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook processed and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
