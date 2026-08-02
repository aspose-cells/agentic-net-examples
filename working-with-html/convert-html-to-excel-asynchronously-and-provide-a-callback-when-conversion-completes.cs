// Title: Async HTML to Excel Conversion with Aspose.Cells and Completion Callback (C#)
// Description: Demonstrates how to load an HTML file into an Aspose.Cells Workbook using LoadOptions, save it as XLSX, run the operation on a background thread with Task.Run, and invoke an optional Action delegate when the conversion finishes.
// Keywords: Aspose.Cells async conversion | HTML to XLSX C# | LoadOptions Html Aspose | Task.Run Excel generation | completion callback C# | background thread file conversion | Aspose.Cells LoadFormat.Html | C# async workbook save | Excel export from HTML | non‑blocking conversion Aspose
// Common Searches: async HTML to Excel conversion Aspose.Cells C# | C# load HTML workbook with Aspose.Cells and save as XLSX | how to run Aspose.Cells conversion on a background thread | execute callback after Aspose.Cells file save | non‑blocking HTML to Excel example .NET
// Developer Intent: Convert an HTML document to an Excel workbook asynchronously and trigger a callback once the file is saved.
// Use Cases: Desktop UI: start conversion on a worker thread and enable UI updates when the XLSX file is ready. | Web API: launch an async HTML‑to‑Excel job and log or notify via a delegate after completion. | Batch processing: convert many HTML reports in parallel without blocking the main service, using a per‑file completion handler.
// AI Prompts: Write C# code that processes a list of HTML files with Aspose.Cells asynchronously, saving each as XLSX and logging the result via a callback. | Show how to add robust exception handling to the async HTML‑to‑Excel conversion while guaranteeing the onCompleted Action runs. | Create a .NET background service that uses HtmlToExcelConverter and raises progress events for each completed conversion.

using System;
using System.Threading.Tasks;
using Aspose.Cells;

// Demonstrates how to load an HTML file into an Aspose.Cells Workbook using LoadOptions, save it as XLSX, run the operation on a background thread with Task.Run, and invoke an optional Action delegate when the conversion finishes.
public class HtmlToExcelConverter
{
    // Asynchronously converts an HTML file to an Excel file and invokes a callback when done.
    public async Task ConvertAsync(string htmlFilePath, string excelFilePath, Action onCompleted)
    {
        await Task.Run(() =>
        {
            // Load the HTML file into a workbook using LoadOptions.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(htmlFilePath, loadOptions);

            // Save the workbook as an Excel file (XLSX format).
            workbook.Save(excelFilePath, SaveFormat.Xlsx);
        });

        // Execute the callback after conversion finishes.
        onCompleted?.Invoke();
    }
}

// Example usage
public class Program
{
    public static async Task Main()
    {
        string htmlPath = "input.html";
        string excelPath = "output.xlsx";

        var converter = new HtmlToExcelConverter();

        await converter.ConvertAsync(htmlPath, excelPath, () =>
        {
            Console.WriteLine("HTML to Excel conversion completed successfully.");
        });
    }
}
