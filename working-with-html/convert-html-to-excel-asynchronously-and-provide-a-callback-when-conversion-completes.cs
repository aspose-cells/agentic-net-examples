// Title: Convert HTML string to XLSX workbook asynchronously with a callback using Aspose.Cells for .NET
// Common Searches: asp.net convert html string to excel file asynchronously with Aspose.Cells | c# async html to xlsx conversion callback Aspose.Cells | how to use HtmlLoadOptions for async HTML to Excel conversion in .NET | run Aspose.Cells HTML to workbook conversion on background thread c# | save workbook to specific path after async conversion Aspose.Cells | handle errors in async Aspose.Cells HTML conversion callback
// Tags: async html to xlsx conversion Aspose.Cells | HtmlLoadOptions memory stream workbook creation | background task workbook generation C# | callback error handling for Aspose.Cells conversion | save workbook to file path using Aspose.Cells

using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;

// The HtmlToExcelConverter class offers a ConvertAsync method that loads an HTML string into a Workbook via HtmlLoadOptions using a MemoryStream, saves it as an XLSX file to a specified path, and invokes a provided callback with any exception, all executed on a background Task.
public class HtmlToExcelConverter
{
    /// <param name="htmlContent">The HTML string to convert.</param>
    /// <param name="outputFilePath">Full path where the Excel file will be saved.</param>
    /// <param name="callback">
    /// Callback invoked after conversion completes.
    /// Parameter is null on success, otherwise contains the exception that occurred.
    /// </param>
    public void ConvertAsync(string htmlContent, string outputFilePath, Action<Exception?> callback)
    {
        // Run the conversion on a background thread.
        Task.Run(() =>
        {
            try
            {
                // Convert HTML string to a workbook using HtmlLoadOptions.
                byte[] htmlBytes = Encoding.UTF8.GetBytes(htmlContent);
                using (var stream = new MemoryStream(htmlBytes))
                {
                    var loadOptions = new HtmlLoadOptions();
                    var workbook = new Workbook(stream, loadOptions);

                    // Ensure the output directory exists.
                    string? directory = Path.GetDirectoryName(outputFilePath);
                    if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    // Save the workbook to the specified file in XLSX format.
                    workbook.Save(outputFilePath, SaveFormat.Xlsx);
                }

                // Invoke callback indicating success.
                callback?.Invoke(null);
            }
            catch (Exception ex)
            {
                // Invoke callback with the caught exception.
                callback?.Invoke(ex);
            }
        });
    }
}

// Entry point for the console application.
public class Program
{
    public static void Main()
    {
        string htmlContent = "<html><body><table><tr><td>Hello</td><td>World</td></tr></table></body></html>";
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Converted.xlsx");

        var converter = new HtmlToExcelConverter();
        converter.ConvertAsync(htmlContent, outputPath, (Exception? ex) =>
        {
            if (ex == null)
            {
                Console.WriteLine($"Conversion succeeded. File saved at: {outputPath}");
            }
            else
            {
                Console.WriteLine($"Conversion failed: {ex.Message}");
            }
        });

        // Keep the console open until the user decides to close it.
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }
}
