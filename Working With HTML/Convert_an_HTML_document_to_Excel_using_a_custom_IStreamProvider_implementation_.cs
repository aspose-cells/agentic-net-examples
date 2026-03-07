using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlToExcel
{
    // Custom stream provider used for loading the HTML file.
    // It opens a read‑only FileStream for the HTML source.
    public class CustomStreamProvider : IStreamProvider
    {
        // Called by Aspose.Cells when it needs a stream.
        public void InitStream(StreamProviderOptions options)
        {
            // options.DefaultPath contains the full path of the HTML file to load.
            // Open the file for reading and assign it to the options.
            options.Stream = File.Open(options.DefaultPath, FileMode.Open, FileAccess.Read);
        }

        // Called after Aspose.Cells finishes using the stream.
        public void CloseStream(StreamProviderOptions options)
        {
            // Ensure the stream is properly closed.
            if (options.Stream != null)
            {
                options.Stream.Close();
            }
        }
    }

    public class HtmlToExcelConverter
    {
        public static void Convert(string htmlFilePath, string excelOutputPath)
        {
            // Configure load options to use the custom stream provider.
            HtmlLoadOptions loadOptions = new HtmlLoadOptions
            {
                StreamProvider = new CustomStreamProvider()
            };

            // Load the HTML document into a workbook using the load options.
            Workbook workbook = new Workbook(htmlFilePath, loadOptions);

            // Save the workbook as XLSX.
            workbook.Save(excelOutputPath, SaveFormat.Xlsx);
        }

        // Example usage.
        public static void Main()
        {
            string htmlPath = "sample.html";      // Path to the source HTML file.
            string excelPath = "result.xlsx";     // Desired XLSX output path.

            Convert(htmlPath, excelPath);

            Console.WriteLine($"HTML file '{htmlPath}' has been converted to Excel file '{excelPath}'.");
        }
    }
}