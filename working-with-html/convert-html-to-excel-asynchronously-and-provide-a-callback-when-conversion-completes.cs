// Title: Convert an HTML string to XLSX asynchronously with Aspose.Cells and receive a callback on completion (C#)
// AI Prompts: Write a C# async method that accepts an HTML string, loads it into an Aspose.Cells Workbook using HtmlLoadOptions, saves the workbook as an XLSX file, and invokes an Action<Exception?> callback after the operation finishes. | Create a console application that calls the asynchronous HTML‑to‑Excel conversion, awaits its completion, and implements a callback method to log success or error information.
// Common Searches: c# aspocells async html to xlsx conversion with callback | how to load html from string into workbook using Aspose.Cells .NET | save workbook as xlsx from memory stream aspnet core async | aspocells HtmlLoadOptions example converting html tables to excel | run html to excel conversion in background task and handle errors c#
// Tags: asynchronous conversion of HTML to XLSX with Aspose.Cells | HtmlLoadOptions for loading HTML from memory stream | callback pattern using Action<Exception?> after async operation | programmatic creation of output directory before workbook save | saving Aspose.Cells Workbook as SaveFormat.Xlsx

using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;

// The example defines an async ConvertAsync method that validates an HTML string, loads it into an Aspose.Cells Workbook via a MemoryStream and HtmlLoadOptions, ensures the output folder exists, saves the workbook as an XLSX file, and calls a provided Action<Exception?> callback to report success (null) or any exception.
public class HtmlToExcelConverter
{
    // Asynchronously converts an HTML string to an Excel file and invokes a callback upon completion.
    public static async Task ConvertAsync(string htmlContent, string outputFilePath, Action<Exception?> callback)
    {
        try
        {
            await Task.Run(() =>
            {
                if (string.IsNullOrEmpty(htmlContent))
                    throw new ArgumentException("HTML content cannot be null or empty.", nameof(htmlContent));

                // Load the HTML content from a memory stream using HtmlLoadOptions.
                var loadOptions = new HtmlLoadOptions();
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(htmlContent)))
                {
                    // Workbook constructor can accept a stream and load options.
                    var workbook = new Workbook(stream, loadOptions);

                    // Ensure the output directory exists.
                    var outputDir = Path.GetDirectoryName(outputFilePath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Save the workbook to the desired Excel format (XLSX in this case).
                    workbook.Save(outputFilePath, SaveFormat.Xlsx);
                }
            });

            // Conversion succeeded; invoke callback with null exception.
            callback?.Invoke(null);
        }
        catch (Exception ex)
        {
            // An error occurred; pass the exception to the callback.
            callback?.Invoke(ex);
        }
    }
}

// Example usage.
class Program
{
    static void Main()
    {
        string html = @"
            <html>
                <body>
                    <table border='1'>
                        <tr><th>Name</th><th>Age</th></tr>
                        <tr><td>Alice</td><td>30</td></tr>
                        <tr><td>Bob</td><td>25</td></tr>
                    </table>
                </body>
            </html>";

        string outputPath = "ConvertedFromHtml.xlsx";

        // Start the asynchronous conversion and wait for it to finish.
        HtmlToExcelConverter.ConvertAsync(html, outputPath, ConversionFinished).Wait();
    }

    // Callback method invoked after conversion completes.
    static void ConversionFinished(Exception? ex)
    {
        if (ex == null)
        {
            Console.WriteLine("HTML successfully converted to Excel.");
        }
        else
        {
            Console.WriteLine($"Conversion failed: {ex.Message}");
        }
    }
}
