// Title: Convert HTML to XLSX with Aspose.Cells using only MemoryStreams (C#)
// Description: A C# example that loads an HTML document into a MemoryStream, creates an Aspose.Cells Workbook from the stream, and saves it to another MemoryStream as an XLSX byte array. The method works entirely in memory, making it ideal for large files, serverless environments, and scenarios where disk I/O must be avoided.
// Keywords: Aspose.Cells | C# | .NET | HTML to Excel | MemoryStream conversion | in‑memory XLSX | large file processing | no temporary files | stream‑based conversion | serverless Excel generation
// Common Searches: Aspose.Cells convert HTML to Excel using MemoryStream | C# convert large HTML file to XLSX without temp files | in‑memory HTML to Excel conversion .NET | load HTML from stream Aspose.Cells | serverless HTML to Excel C# example
// Developer Intent: The developer needs to transform an HTML document into an Excel workbook entirely in memory to eliminate disk I/O, improve performance for large inputs, and fit environments that restrict file system access.
// Use Cases: Expose a web API that receives an HTML report, converts it to XLSX on‑the‑fly, and streams the result back to the client. | Run a background service that batch‑processes massive HTML logs, generating Excel files without creating intermediate files on disk. | Deploy a serverless function (e.g., Azure Functions or AWS Lambda) that converts incoming HTML payloads to Excel using only memory resources.
// AI Prompts: Write C# code that uses Aspose.Cells to read HTML from a MemoryStream and return the workbook as an XLSX byte array with comprehensive error handling. | Show how to refactor the conversion method to accept any Stream (e.g., network request body) and output the Excel data as a Stream instead of a byte array. | Explain how to extend the sample to save the workbook in additional formats such as CSV, ODS, or PDF while keeping the entire process in memory.

using System;
using System.IO;
using Aspose.Cells;

// A C# example that loads an HTML document into a MemoryStream, creates an Aspose.Cells Workbook from the stream, and saves it to another MemoryStream as an XLSX byte array. The method works entirely in memory, making it ideal for large files, serverless environments, and scenarios where disk I/O must be avoided.
public class HtmlToExcelConverter
{
    // Converts an HTML file to an Excel file using only memory streams.
    // Returns the Excel file as a byte array (XLSX format).
    public static byte[] ConvertHtmlToExcel(string htmlFilePath)
    {
        // Verify that the source HTML file exists.
        if (!File.Exists(htmlFilePath))
            throw new FileNotFoundException($"HTML file not found: {htmlFilePath}");

        try
        {
            // Load the HTML content into a memory stream (no temporary files on disk).
            using (FileStream fileStream = new FileStream(htmlFilePath, FileMode.Open, FileAccess.Read))
            using (MemoryStream htmlStream = new MemoryStream())
            {
                fileStream.CopyTo(htmlStream);
                htmlStream.Position = 0; // Reset for reading.

                // Load the workbook from the HTML stream.
                Workbook workbook = new Workbook(htmlStream);

                // Save the workbook to another memory stream in XLSX format.
                using (MemoryStream excelStream = new MemoryStream())
                {
                    workbook.Save(excelStream, SaveFormat.Xlsx);
                    return excelStream.ToArray();
                }
            }
        }
        catch (Exception ex)
        {
            // Wrap and rethrow to provide context while preserving stack trace.
            throw new InvalidOperationException("Failed to convert HTML to Excel.", ex);
        }
    }

    // Example usage.
    public static void Main()
    {
        string htmlPath = "large_input.html"; // Path to the source HTML file.

        try
        {
            byte[] excelBytes = ConvertHtmlToExcel(htmlPath);

            // Optionally write the result to a file for verification.
            File.WriteAllBytes("converted_output.xlsx", excelBytes);
            Console.WriteLine($"HTML successfully converted to Excel. Output size: {excelBytes.Length} bytes.");
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.WriteLine(fnfEx.Message);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
