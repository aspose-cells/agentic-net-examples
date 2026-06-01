using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsHtmlToExcel
{
    public class Converter
    {
        /// <summary>
        /// Converts an HTML document to an Excel workbook using memory streams only.
        /// </summary>
        /// <param name="htmlInputPath">Path to the source HTML file.</param>
        /// <param name="excelOutputPath">Path where the resulting Excel file will be saved.</param>
        public static void ConvertHtmlToExcel(string htmlInputPath, string excelOutputPath)
        {
            // Open the HTML file as a read‑only memory stream (no temporary files on disk)
            using (MemoryStream htmlStream = new MemoryStream())
            {
                using (FileStream fileStream = File.OpenRead(htmlInputPath))
                {
                    fileStream.CopyTo(htmlStream);
                }

                // Reset the position so the workbook can read from the beginning
                htmlStream.Position = 0;

                // Load options specifying that the source format is HTML
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);

                // Load the HTML content into a workbook (uses the stream + load options)
                Workbook workbook = new Workbook(htmlStream, loadOptions);

                // Prepare an output memory stream for the Excel file
                using (MemoryStream excelStream = new MemoryStream())
                {
                    // Save the workbook to the memory stream in XLSX format
                    // (uses the provided Save(Stream, SaveFormat) rule)
                    workbook.Save(excelStream, SaveFormat.Xlsx);

                    // Reset the stream position before writing it out
                    excelStream.Position = 0;

                    // Write the Excel bytes to the final file location
                    using (FileStream outFile = File.Create(excelOutputPath))
                    {
                        excelStream.CopyTo(outFile);
                    }
                }
            }
        }

        // Example usage
        public static void Main()
        {
            string htmlPath = "sample.html";      // source HTML file
            string excelPath = "result.xlsx";     // destination Excel file

            try
            {
                ConvertHtmlToExcel(htmlPath, excelPath);
                Console.WriteLine($"HTML successfully converted to Excel: {excelPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Conversion failed: {ex.Message}");
            }
        }
    }
}