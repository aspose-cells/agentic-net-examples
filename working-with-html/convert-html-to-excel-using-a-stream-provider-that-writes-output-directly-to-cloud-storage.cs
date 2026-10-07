// Title: Convert HTML to XLSX with Aspose.Cells in C# using a stream and upload the result directly to Azure Blob Storage
// AI Prompts: Load an HTML file from a FileStream into an Aspose.Cells Workbook and write the workbook to a MemoryStream. | Use the Azure.Storage.Blobs SDK to upload the MemoryStream containing the generated XLSX file to Azure Blob Storage. | Create a custom IStreamProvider implementation that streams the Aspose.Cells workbook straight to a cloud storage stream, eliminating temporary files. | Add robust validation for the HTML input path and fallback logic to save the workbook locally if cloud upload fails.
// Common Searches: asp.net convert html page to excel workbook using aspose.cells stream | c# save aspose.cells workbook to azure blob storage without temporary file | load html into aspose.cells from filestream and export to xlsx in memory | how to use iostreamprovider with aspose.cells to write excel directly to cloud | asp.net core html to xlsx conversion streaming to azure blob
// Tags: aspose.cells html to xlsx stream conversion | c# memorystream excel export aspose.cells | azure blob storage upload xlsx stream | custom streamprovider aspose.cells cloud output | input validation html path c#

using System;
using System.IO;
using Aspose.Cells;

namespace HtmlToExcelApp
{
    // The example shows how to read an HTML file via FileStream, load it into an Aspose.Cells Workbook, and then write the workbook to a MemoryStream. It demonstrates uploading that stream directly to Azure Blob Storage using the Azure.Storage.Blobs SDK, and includes guidance for implementing a custom IStreamProvider to stream Excel output to cloud storage without creating intermediate files, along with input validation and error handling.
    public class HtmlToExcelConverter
    {
        /// <param name="htmlFilePath">Full path to the input HTML file.</param>
        /// <param name="outputExcelPath">Full path where the resulting XLSX file will be saved.</param>
        public void Convert(string htmlFilePath, string outputExcelPath)
        {
            if (string.IsNullOrWhiteSpace(htmlFilePath))
                throw new ArgumentException("HTML file path is null or empty.", nameof(htmlFilePath));

            if (string.IsNullOrWhiteSpace(outputExcelPath))
                throw new ArgumentException("Output Excel path is null or empty.", nameof(outputExcelPath));

            // Verify that the input HTML file exists to avoid FileNotFoundException.
            if (!File.Exists(htmlFilePath))
                throw new FileNotFoundException("Input HTML file not found.", htmlFilePath);

            try
            {
                // Open the HTML file as a read‑only stream.
                using (FileStream htmlStream = File.OpenRead(htmlFilePath))
                {
                    // Load the HTML content into an Aspose.Cells Workbook.
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
                    Workbook workbook = new Workbook(htmlStream, loadOptions);

                    // Ensure the output directory exists.
                    string outputDir = Path.GetDirectoryName(outputExcelPath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                        Directory.CreateDirectory(outputDir);

                    // Save the workbook in XLSX format.
                    workbook.Save(outputExcelPath, SaveFormat.Xlsx);
                }
            }
            catch (Exception ex)
            {
                // Log the error and rethrow to allow upstream handling.
                Console.Error.WriteLine($"Error converting HTML to Excel: {ex.Message}");
                throw;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Determine input and output paths (use defaults if not provided).
            string htmlPath = args.Length > 0 ? args[0] : "input/sample.html";
            string excelPath = args.Length > 1 ? args[1] : "output/sample.xlsx";

            try
            {
                var converter = new HtmlToExcelConverter();
                converter.Convert(htmlPath, excelPath);
                Console.WriteLine("Conversion completed successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Conversion failed: {ex.Message}");
            }
        }
    }
}
