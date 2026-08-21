// Title: Async Conversion of WordArt‑Rich Excel Sheets to HTML with Aspose.Cells for .NET
// Description: This C# example shows how to load an Excel workbook that contains WordArt shapes with gradient fills, set HtmlSaveOptions for HTML5 with separate CSS and base64‑encoded images, and convert each worksheet’s used range to its own HTML file on separate threads. It creates an output folder, runs a Task per sheet, writes the HTML bytes, and awaits all tasks, delivering fast, non‑blocking conversion of WordArt‑enabled spreadsheets.
// Keywords: Aspose.Cells | C# | .NET | asynchronous HTML conversion | Excel to HTML | WordArt gradients | parallel worksheet conversion | HtmlSaveOptions | Range.ToHtml | Task.Run | Task.WhenAll | HTML5 export | base64 images
// Common Searches: async convert Excel to HTML Aspose.Cells | export WordArt gradients to HTML .NET | parallel worksheet HTML conversion C# | how to use Range.ToHtml asynchronously | Aspose.Cells generate HTML5 with separate CSS
// Developer Intent: The developer wants to transform every worksheet that contains WordArt with gradient fills into separate HTML files, executing the conversions concurrently to avoid blocking the application.
// Use Cases: Generate web‑ready HTML reports from Excel workbooks that include WordArt graphics without UI delays. | Batch‑process large Excel files on a server, converting each sheet to HTML in parallel to reduce overall runtime. | Build an API endpoint that accepts an uploaded spreadsheet and returns per‑sheet HTML files, preserving gradient styling via base64 images. | Create automated documentation pipelines that convert design‑heavy Excel sheets to HTML for publishing.
// AI Prompts: Modify the example to write conversion errors to a log file while preserving asynchronous processing. | Replace the Task.Run loop with Parallel.ForEach and explain the differences. | Show how to configure HtmlSaveOptions to output WordArt gradients as SVG instead of base64 images. | Add cancellation support using CancellationToken to stop the conversion mid‑process. | Provide code to stream the generated HTML directly to an HTTP response in an ASP.NET Core controller.

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This C# example shows how to load an Excel workbook that contains WordArt shapes with gradient fills, set HtmlSaveOptions for HTML5 with separate CSS and base64‑encoded images, and convert each worksheet’s used range to its own HTML file on separate threads. It creates an output folder, runs a Task per sheet, writes the HTML bytes, and awaits all tasks, delivering fast, non‑blocking conversion of WordArt‑enabled spreadsheets.
public class AsyncWordArtToHtmlConverter
{
    public static async Task Main(string[] args)
    {
        // Path to the source Excel file (contains WordArt with gradients)
        string sourcePath = "WordArtWorkbook.xlsx";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Error: The file '{sourcePath}' was not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(sourcePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            HtmlVersion = HtmlVersion.Html5,
            ExportWorksheetCSSSeparately = true,
            ExportImagesAsBase64 = true
        };

        // Prepare output folder
        string outputFolder = "HtmlOutput";
        Directory.CreateDirectory(outputFolder);

        // Create a task for each worksheet to convert its used range to HTML
        Task[] conversionTasks = new Task[workbook.Worksheets.Count];
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            int sheetIndex = i; // capture loop variable
            conversionTasks[i] = Task.Run(() =>
            {
                try
                {
                    Worksheet sheet = workbook.Worksheets[sheetIndex];

                    // Determine the used range of the worksheet
                    Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                    // Convert the range to HTML bytes
                    byte[] htmlBytes = usedRange.ToHtml(htmlOptions);

                    // Build the output file name
                    string htmlFilePath = Path.Combine(outputFolder, $"Sheet{sheetIndex + 1}.html");

                    // Write the HTML bytes to file
                    File.WriteAllBytes(htmlFilePath, htmlBytes);

                    Console.WriteLine($"Worksheet '{sheet.Name}' converted to HTML: {htmlFilePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing worksheet index {sheetIndex}: {ex.Message}");
                }
            });
        }

        // Await all conversion tasks to complete
        await Task.WhenAll(conversionTasks);

        Console.WriteLine("All worksheets have been converted to HTML asynchronously.");
    }
}
