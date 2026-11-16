// Title: Convert an HTML stream to an Excel .xlsx workbook and save it to a file path using Aspose.Cells in C#
// Common Searches: how to load HTML from a Stream into an Aspose.Cells Workbook in C# | saving an Aspose.Cells workbook to a specific folder without temporary files | using HtmlLoadOptions to convert HTML tables to Excel with Aspose.Cells | C# code to create missing directories before saving an Excel file | convert HTML string to MemoryStream and generate XLSX using Aspose.Cells | error handling when converting HTML to Excel with Aspose.Cells
// Tags: Aspose.Cells HTML stream to XLSX conversion | C# save workbook to file path Aspose.Cells | HtmlLoadOptions load HTML into workbook | stream-based Excel generation without temporary files | directory creation before workbook save C# | exception handling for HTML to Excel conversion

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace HtmlToExcelConverter
{
    // Demonstrates loading HTML content from a .NET Stream into an Aspose.Cells Workbook using HtmlLoadOptions, then saving the workbook as an .xlsx file to a specified location, including automatic directory creation and basic error handling.
    public class Converter
    {
        /// <param name="htmlStream">Input stream containing the HTML content.</param>
        /// <param name="connectionString">Azure Storage account connection string (not used in this implementation).</param>
        /// <param name="containerName">Name of the blob container where the Excel file would be stored (not used in this implementation).</param>
        /// <param name="blobName">Name of the blob (including .xlsx extension). Used as the local file name.</param>
        public void ConvertHtmlToExcel(Stream htmlStream, string connectionString, string containerName, string blobName)
        {
            if (htmlStream == null) throw new ArgumentNullException(nameof(htmlStream));
            if (string.IsNullOrWhiteSpace(blobName)) throw new ArgumentException("Blob name must be provided.", nameof(blobName));

            try
            {
                // Load HTML content into an Aspose.Cells Workbook.
                var workbook = new Workbook(htmlStream, new HtmlLoadOptions());

                // Determine output path (current directory + blobName).
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), blobName);

                // Ensure the target directory exists.
                string directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Save the workbook as XLSX to the local file system.
                workbook.Save(outputPath, SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                // Log the error and rethrow for upstream handling.
                Console.Error.WriteLine($"Error converting HTML to Excel: {ex.Message}");
                throw;
            }
        }
    }

    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                // Sample HTML content.
                string html = "<html><body><table><tr><td>Sample</td></tr></table></body></html>";
                using var htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(html));

                // Output Excel file name.
                string outputFileName = "output.xlsx";

                // Remove existing file to avoid conflicts.
                if (File.Exists(outputFileName))
                {
                    File.Delete(outputFileName);
                }

                // Perform conversion.
                var converter = new Converter();
                converter.ConvertHtmlToExcel(htmlStream, string.Empty, string.Empty, outputFileName);

                Console.WriteLine($"Excel file successfully saved to: {Path.GetFullPath(outputFileName)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex}");
            }
        }
    }
}
