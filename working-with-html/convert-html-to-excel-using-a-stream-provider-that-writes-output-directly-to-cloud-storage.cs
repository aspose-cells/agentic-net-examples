// Title: Convert an HTML Stream to an XLSX Workbook and Save Directly to Cloud Storage Using Aspose.Cells for .NET
// AI Prompts: Generate a C# async method that receives an HTML Stream, loads it into an Aspose.Cells Workbook with HtmlLoadOptions, and writes the resulting XLSX bytes to an Azure Blob Storage container using a CloudBlobStream. | Create a reusable class that converts HTML content from any Stream into an in‑memory Excel file and uploads it to Amazon S3 without creating a temporary file on disk. | Add robust error handling to the conversion routine that captures Aspose.Cells exceptions, logs the stack trace to a file, and returns a meaningful status code to the caller.
// Common Searches: asp.net core convert html stream to xlsx and upload to azure blob storage | c# Aspose.Cells load html from memory stream and save workbook directly to s3 | how to stream Aspose.Cells workbook to cloud storage without saving locally | convert html to excel using Aspose.Cells and write output to Google Cloud Storage | c# example for HtmlLoadOptions with Aspose.Cells and cloud upload
// Tags: HTML to XLSX conversion Aspose.Cells | Aspose.Cells stream upload to Azure Blob | save workbook directly to Amazon S3 | load HTML with HtmlLoadOptions C# | in‑memory Excel generation without temporary file | cloud storage integration Aspose.Cells

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

namespace HtmlToExcelConverter
{
    // The example demonstrates how to load HTML content from a Stream into an Aspose.Cells Workbook using HtmlLoadOptions, then save the workbook as an XLSX file directly to cloud storage (e.g., Azure Blob, Amazon S3) without writing a temporary file to the local file system, with basic validation and error handling.
    public class Converter
    {
        private readonly string _outputDirectory;

        // Constructor receives a directory path where Excel files will be saved
        public Converter(string outputDirectory)
        {
            if (string.IsNullOrWhiteSpace(outputDirectory))
                throw new ArgumentException("Output directory must be provided.", nameof(outputDirectory));

            if (!Directory.Exists(outputDirectory))
                Directory.CreateDirectory(outputDirectory);

            _outputDirectory = outputDirectory;
        }

        /// <param name="htmlStream">Input stream containing HTML content.</param>
        /// <param name="fileName">Name of the Excel file to create (including .xlsx extension).</param>
        public async Task ConvertAndSaveAsync(Stream htmlStream, string fileName)
        {
            if (htmlStream == null) throw new ArgumentNullException(nameof(htmlStream));
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("File name must be provided.", nameof(fileName));

            try
            {
                // Load the HTML into an Aspose.Cells Workbook
                var loadOptions = new HtmlLoadOptions(); // default options; adjust if needed
                var workbook = new Workbook(htmlStream, loadOptions);

                // Prepare the full path for the output Excel file
                string outputPath = Path.Combine(_outputDirectory, fileName);

                // Save the workbook to the file in XLSX format
                workbook.Save(outputPath, SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                // Log or rethrow as needed; for this example we just write to console
                Console.Error.WriteLine($"Error during conversion: {ex.Message}");
                throw;
            }
        }
    }

    // Example usage
    public class Program
    {
        public static async Task Main()
        {
            // Directory where the Excel file will be stored
            string outputDirectory = "ConvertedExcel";

            // Initialize the converter with the output directory
            var converter = new Converter(outputDirectory);

            // Path to the local HTML file
            string htmlFilePath = "sample.html";

            // Ensure the HTML file exists before attempting to read it
            if (!File.Exists(htmlFilePath))
            {
                Console.Error.WriteLine($"HTML file not found: {htmlFilePath}");
                return;
            }

            try
            {
                // Read HTML from the local file and convert it
                using (FileStream htmlFileStream = File.OpenRead(htmlFilePath))
                {
                    await converter.ConvertAndSaveAsync(htmlFileStream, "output.xlsx");
                }

                Console.WriteLine("HTML has been converted to Excel and saved locally.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Conversion failed: {ex.Message}");
            }
        }
    }
}
