// Title: Detect Encrypted Excel Workbooks in a SharePoint Library Using Aspose.Cells (C#)
// Description: A C# console app that iterates over SharePoint Excel file URLs, downloads each workbook with HttpClient, uses Aspose.Cells FileFormatUtil to identify the file format and encryption status, logs the results for compliance, and optionally saves encrypted files locally for further analysis.
// Keywords: Aspose.Cells | C# | SharePoint | Excel encryption detection | FileFormatUtil | IsEncrypted | Office 365 compliance | download Excel from SharePoint | batch workbook audit | encrypted workbook logging
// Common Searches: how to check if SharePoint Excel files are password protected using Aspose.Cells | C# code to detect encrypted workbooks in a SharePoint library | Aspose.Cells detect encrypted Excel file | automate Excel encryption audit in SharePoint | download and scan Excel files for encryption with .NET
// Developer Intent: Programmatically identify which Excel workbooks stored in a SharePoint document library are encrypted and record their status to satisfy security and compliance requirements.
// Use Cases: Run a scheduled compliance scan that flags encrypted Excel files across a SharePoint site and generates a summary report. | Batch download all encrypted workbooks for secure archiving, decryption, or further forensic analysis. | Integrate the detection logic into a larger governance workflow that logs format and encryption details for each file.
// AI Prompts: Generate a C# method that accepts a collection of SharePoint file URLs, downloads each workbook, and returns a dictionary of URL → IsEncrypted using Aspose.Cells. | Enhance the sample to capture the workbook's password hint (if available) and include it in the compliance log. | Create a PowerShell script that runs the compiled .NET executable, parses its console output, and writes the encryption results to a CSV file for reporting.

using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

namespace SharePointWorkbookEncryptionCheck
{
    // A C# console app that iterates over SharePoint Excel file URLs, downloads each workbook with HttpClient, uses Aspose.Cells FileFormatUtil to identify the file format and encryption status, logs the results for compliance, and optionally saves encrypted files locally for further analysis.
    class Program
    {
        // Entry point
        static async Task Main(string[] args)
        {
            // List of SharePoint file URLs to inspect.
            // Replace these with actual URLs from your SharePoint library.
            List<string> workbookUrls = new List<string>
            {
                "https://sharepoint.example.com/sites/Docs/Workbook1.xlsx",
                "https://sharepoint.example.com/sites/Docs/Workbook2.xls",
                // Add more URLs as needed
            };

            // HttpClient instance for downloading files.
            using (HttpClient httpClient = new HttpClient())
            {
                foreach (string fileUrl in workbookUrls)
                {
                    try
                    {
                        // Download the workbook into a memory stream.
                        using (Stream fileStream = await httpClient.GetStreamAsync(fileUrl))
                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            await fileStream.CopyToAsync(memoryStream);
                            memoryStream.Position = 0; // Reset stream position for detection.

                            // Detect file format and encryption status using Aspose.Cells.
                            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(memoryStream);
                            bool isEncrypted = formatInfo.IsEncrypted;

                            // Log the result.
                            Console.WriteLine($"File: {fileUrl}");
                            Console.WriteLine($"  Encrypted: {isEncrypted}");
                            Console.WriteLine($"  Format: {formatInfo.FileFormatType}");
                            Console.WriteLine();

                            // Optional: Save the file locally for further analysis.
                            if (isEncrypted)
                            {
                                // Example: Save encrypted file to a local folder.
                                string localFileName = Path.GetFileName(new Uri(fileUrl).LocalPath);
                                string localPath = Path.Combine("DownloadedWorkbooks", localFileName);
                                Directory.CreateDirectory(Path.GetDirectoryName(localPath));

                                // Rewind the stream before saving.
                                memoryStream.Position = 0;
                                using (FileStream localFile = new FileStream(localPath, FileMode.Create, FileAccess.Write))
                                {
                                    await memoryStream.CopyToAsync(localFile);
                                }

                                Console.WriteLine($"  Encrypted workbook saved to: {localPath}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log any errors encountered while processing the file.
                        Console.WriteLine($"Error processing file '{fileUrl}': {ex.Message}");
                    }
                }
            }
        }
    }
}
