// Title: Convert HTML to XLSX asynchronously with Aspose.Cells in C# and receive a completion callback
// AI Prompts: Implement an async C# routine that receives paths for an HTML source and an XLSX target, creates a LoadOptions object for HTML, opens the file with Aspose.Cells Workbook, saves it as Xlsx, and triggers a supplied Action delegate once the operation finishes. | Demonstrate invoking the async routine from Main, checking that the HTML file exists, handling exceptions, and passing a lambda that prints a completion message.
// Common Searches: aspocells async html to xlsx conversion with completion handler c# | c# load html file into Aspose.Cells workbook on background thread | example using Action callback after Aspose.Cells save operation | how to wrap Aspose.Cells HTML import in Task.Run | convert html to excel asynchronously in a .NET console application
// Tags: Aspose.Cells HTML import to XLSX async | C# Action callback after workbook save | LoadOptions LoadFormat.Html usage Aspose.Cells | Task.Run pattern for Aspose.Cells operations | saving workbook as Xlsx with Aspose.Cells

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

namespace HtmlToExcelAsyncDemo
{
    // The example defines ConvertHtmlToExcelAsync, which validates the HTML file, loads it into an Aspose.Cells Workbook using LoadOptions for HTML, saves the workbook as an XLSX file, and then invokes an optional Action callback. The Main method shows how to call this async method with file existence checks, exception handling, and a completion message.
    public static class Converter
    {
        /// <param name="htmlFilePath">Full path to the source HTML file.</param>
        /// <param name="excelFilePath">Full path where the resulting Excel file will be saved.</param>
        /// <param name="onCompleted">Callback invoked after successful conversion.</param>
        /// <returns>A task representing the asynchronous operation.</returns>
        public static async Task ConvertHtmlToExcelAsync(string htmlFilePath, string excelFilePath, Action onCompleted)
        {
            await Task.Run(() =>
            {
                // Verify the source HTML file exists.
                if (!File.Exists(htmlFilePath))
                    throw new FileNotFoundException("HTML source file not found.", htmlFilePath);

                // Load the HTML content into a workbook using HTML load options.
                var loadOptions = new LoadOptions(LoadFormat.Html);
                var workbook = new Workbook(htmlFilePath, loadOptions);

                // Save the workbook in XLSX format.
                workbook.Save(excelFilePath, SaveFormat.Xlsx);
            });

            // Invoke the callback after the conversion completes.
            onCompleted?.Invoke();
        }
    }

    class Program
    {
        static async Task Main(string[] args)
        {
            // Example usage: adjust paths as needed.
            string htmlPath = "input.html";
            string excelPath = "output.xlsx";

            // Ensure the HTML file exists before attempting conversion.
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: The file '{htmlPath}' was not found.");
                return;
            }

            try
            {
                await Converter.ConvertHtmlToExcelAsync(htmlPath, excelPath, () =>
                {
                    Console.WriteLine("Conversion completed.");
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Conversion failed: {ex.Message}");
            }
        }
    }
}
