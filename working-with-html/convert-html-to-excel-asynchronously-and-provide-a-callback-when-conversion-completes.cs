// Title: Async HTML‑to‑Excel conversion with callback using Aspose.Cells in C#
// Description: Demonstrates how to load an HTML file into an Aspose.Cells Workbook with LoadOptions, save it as an XLSX file, and run the whole process on a background thread. A caller‑supplied Action<bool> callback reports success or failure, and the sample includes basic file‑existence checks and a console demo.
// Keywords: Aspose.Cells | HTML to XLSX | C# async conversion | callback after conversion | LoadOptions Html | Task.Run | background file conversion | Excel export .NET | ConvertAsync method | error handling
// Common Searches: how to convert html to excel asynchronously c# | aspocells html to xlsx with callback | c# async html to xlsx conversion example | using task.run to export html as excel | aspocells load html workbook async
// Developer Intent: Create a non‑blocking HTML‑to‑Excel conversion that notifies the caller when the operation finishes.
// Use Cases: Process a batch of HTML reports on a worker thread and update a UI status label via the callback. | Expose a web API that receives HTML, triggers ConvertAsync, and returns a response once the XLSX file is ready. | Schedule a nightly job that converts generated HTML dashboards to Excel and logs the result through the callback.
// AI Prompts: Write a unit test for ConvertAsync that asserts the callback receives true for a valid HTML file and false for a missing file. | Refactor ConvertAsync to return Task<bool> instead of using Action<bool> and show how to await the method. | Add structured exception logging inside the catch block and modify the callback to provide error details to the caller.

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// Demonstrates how to load an HTML file into an Aspose.Cells Workbook with LoadOptions, save it as an XLSX file, and run the whole process on a background thread. A caller‑supplied Action<bool> callback reports success or failure, and the sample includes basic file‑existence checks and a console demo.
public class HtmlToExcelConverter
{
    // Asynchronously converts an HTML file to an Excel file.
    // callback is invoked with true on success, false on failure.
    public static void ConvertAsync(string htmlFilePath, string excelFilePath, Action<bool> callback)
    {
        Task.Run(() =>
        {
            try
            {
                // Verify that the source HTML file exists.
                if (!File.Exists(htmlFilePath))
                    throw new FileNotFoundException("HTML file not found.", htmlFilePath);

                // Load the HTML file into a workbook using LoadOptions.
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
                Workbook workbook = new Workbook(htmlFilePath, loadOptions);

                // Save the workbook as an Excel file (XLSX format).
                workbook.Save(excelFilePath, SaveFormat.Xlsx);

                // Invoke the callback indicating success.
                callback?.Invoke(true);
            }
            catch
            {
                // In case of any error, invoke the callback indicating failure.
                callback?.Invoke(false);
            }
        });
    }
}

public class Program
{
    // Entry point required for compilation.
    public static void Main(string[] args)
    {
        // Example usage:
        string inputHtml = "input.html";
        string outputXlsx = "output.xlsx";

        // Ensure the input file exists before attempting conversion.
        if (!File.Exists(inputHtml))
        {
            Console.WriteLine($"Error: The file '{inputHtml}' does not exist.");
            return;
        }

        HtmlToExcelConverter.ConvertAsync(inputHtml, outputXlsx, success =>
        {
            Console.WriteLine(success ? "Conversion completed." : "Conversion failed.");
        });

        // Prevent the console from closing immediately.
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }
}
