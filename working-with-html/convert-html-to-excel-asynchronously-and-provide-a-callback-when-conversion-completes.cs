using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

namespace HtmlToExcelApp
{
    public class HtmlToExcelConverter
    {
        // Asynchronously converts an HTML file to an Excel file.
        // The callback is invoked with true if conversion succeeds, false otherwise.
        public static async Task ConvertHtmlToExcelAsync(string htmlFilePath, string excelFilePath, Action<bool> callback)
        {
            bool success = false;
            try
            {
                // Verify that the source HTML file exists to avoid FileNotFoundException.
                if (!File.Exists(htmlFilePath))
                    throw new FileNotFoundException($"HTML file not found: {htmlFilePath}");

                await Task.Run(() =>
                {
                    // Load the HTML file into a workbook using LoadOptions to specify the format.
                    var loadOptions = new LoadOptions(LoadFormat.Html);
                    var workbook = new Workbook(htmlFilePath, loadOptions);

                    // Save the workbook as an Excel file (XLSX format).
                    workbook.Save(excelFilePath, SaveFormat.Xlsx);
                });

                success = true;
            }
            catch (Exception ex)
            {
                // Log or handle the exception as needed.
                Console.Error.WriteLine($"Conversion error: {ex.Message}");
                success = false;
            }
            finally
            {
                // Invoke the callback with the result of the conversion.
                callback?.Invoke(success);
            }
        }
    }

    class Program
    {
        // Entry point for the application.
        static async Task Main(string[] args)
        {
            string htmlPath = args.Length > 0 ? args[0] : "input.html";
            string excelPath = args.Length > 1 ? args[1] : "output.xlsx";

            try
            {
                await HtmlToExcelConverter.ConvertHtmlToExcelAsync(htmlPath, excelPath, result =>
                {
                    Console.WriteLine(result ? "Conversion completed successfully." : "Conversion failed.");
                });
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}