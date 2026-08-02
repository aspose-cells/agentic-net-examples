// Title: In‑Memory HTML to Excel Conversion with Aspose.Cells for .NET (C#)
// Description: C# example that loads an HTML file via FileStream, converts it to an XLSX workbook using Aspose.Cells, and returns the result in a MemoryStream—eliminating temporary files and enabling efficient processing of large HTML documents.
// Keywords: Aspose.Cells HTML to XLSX | C# in‑memory conversion | load HTML workbook from stream | save workbook to MemoryStream | large HTML file Excel export | .NET no temp files conversion | stream‑based HTML to Excel
// Common Searches: convert html to excel without temporary files c# | aspacells load html from filestream | memorystream excel output aspnet core | large html report to xlsx in memory | c# aspocells html to xlsx streaming
// Developer Intent: Create an XLSX workbook from an HTML source entirely in memory to avoid disk I/O.
// Use Cases: Process massive HTML reports on a server without consuming disk space. | Return the generated Excel file directly from an ASP.NET Core API endpoint. | Batch‑convert multiple HTML files in a background service using a stream‑only workflow.
// AI Prompts: Generate C# code that reads an HTML file from a FileStream, converts it to XLSX with Aspose.Cells, and outputs a MemoryStream ready for download. | Explain how to adapt the method to accept any input Stream (e.g., network or blob storage) instead of a file path. | Show how to integrate this in‑memory conversion into an ASP.NET Core controller that streams the Excel file as a FileResult.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlToExcel
{
    // C# example that loads an HTML file via FileStream, converts it to an XLSX workbook using Aspose.Cells, and returns the result in a MemoryStream—eliminating temporary files and enabling efficient processing of large HTML documents.
    public static class HtmlToExcelConverter
    {
        /// <param name="htmlFilePath">Full path of the source HTML file.</param>
        /// <returns>A MemoryStream containing the generated Excel file (XLSX format).</returns>
        public static MemoryStream Convert(string htmlFilePath)
        {
            // Validate input
            if (string.IsNullOrEmpty(htmlFilePath))
                throw new ArgumentException("HTML file path must be provided.", nameof(htmlFilePath));

            if (!File.Exists(htmlFilePath))
                throw new FileNotFoundException("HTML file not found.", htmlFilePath);

            // Open the HTML file as a read‑only stream (no temporary files are created)
            using (FileStream htmlStream = new FileStream(htmlFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                // LoadOptions tell Aspose.Cells that the source format is HTML
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);

                // Load the HTML content into a Workbook instance
                Workbook workbook = new Workbook(htmlStream, loadOptions);

                // Prepare a memory stream that will hold the resulting Excel file
                MemoryStream excelStream = new MemoryStream();

                // Save the workbook to the memory stream in XLSX format
                workbook.Save(excelStream, SaveFormat.Xlsx);

                // Reset the position so that callers can read from the beginning
                excelStream.Position = 0;

                // Return the stream (caller is responsible for disposing it)
                return excelStream;
            }
        }

        // Example usage
        public static void RunExample()
        {
            string htmlPath = "large_input.html";

            try
            {
                // Convert HTML to Excel (result is in memory)
                using (MemoryStream excelData = Convert(htmlPath))
                {
                    // For demonstration, write the stream to a file (optional)
                    using (FileStream file = new FileStream("converted.xlsx", FileMode.Create, FileAccess.Write))
                    {
                        excelData.CopyTo(file);
                    }

                    Console.WriteLine("HTML successfully converted to Excel (in-memory).");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    internal class Program
    {
        private static void Main(string[] args)
        {
            HtmlToExcelConverter.RunExample();
        }
    }
}
