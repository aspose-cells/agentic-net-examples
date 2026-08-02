// Title: Async Convert WordArt‑Rich Excel Sheets to HTML with Base64 Gradients – Aspose.Cells for .NET
// Description: Loads an Excel workbook containing WordArt and gradient shapes, then creates a separate HTML file for each worksheet. The code uses HtmlSaveOptions to embed all shape images as Base64, leverages Range.ToHtml for rendering, and runs each sheet conversion on its own Task, awaiting all tasks before finishing.
// Keywords: Aspose.Cells async HTML export | WordArt gradient to Base64 | parallel worksheet conversion .NET | Range.ToHtml example | C# Excel to HTML asynchronous
// Common Searches: convert Excel with WordArt to HTML asynchronously | embed gradient shapes as Base64 when saving Excel as HTML | parallel sheet to HTML Aspose.Cells C# | async Range.ToHtml usage
// Developer Intent: Generate per‑worksheet HTML output from a WordArt‑filled workbook without temporary image files, using asynchronous tasks.
// Use Cases: Create fast HTML previews of every sheet in a marketing report that contains styled WordArt logos. | Expose an ASP.NET Core endpoint that receives an Excel file and returns HTML streams for each sheet, reducing latency with parallel processing. | Batch‑process a directory of Excel templates with decorative shapes, archiving them as self‑contained HTML pages.
// AI Prompts: Rewrite the conversion logic with Parallel.ForEach and ensure the Workbook is disposed correctly. | Show how to stream the generated HTML directly from a MemoryStream in an ASP.NET Core controller. | Add Serilog logging to capture errors per worksheet, including sheet name and exception details.

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// Loads an Excel workbook containing WordArt and gradient shapes, then creates a separate HTML file for each worksheet. The code uses HtmlSaveOptions to embed all shape images as Base64, leverages Range.ToHtml for rendering, and runs each sheet conversion on its own Task, awaiting all tasks before finishing.
class AsyncWordArtHtmlConversion
{
    static async Task Main(string[] args)
    {
        // Path to the source workbook that contains WordArt and gradient shapes
        string sourcePath = "WordArtWorkbook.xlsx";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Error: The source workbook \"{sourcePath}\" was not found.");
            return;
        }

        // Directory where individual HTML files will be written
        string outputDir = "HtmlOutput";
        Directory.CreateDirectory(outputDir);

        Workbook workbook;
        try
        {
            // Load the workbook (uses Aspose.Cells load rule)
            workbook = new Workbook(sourcePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Configure HTML save options – export images (including WordArt gradients) as Base64
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportImagesAsBase64 = true,               // embed shape images directly
            HtmlCrossStringType = HtmlCrossType.Cross, // faster cross‑string handling for large files
            ExportWorksheetCSSSeparately = true        // optional: separate CSS per sheet
        };

        // Prepare a task for each worksheet to run conversion in parallel
        Task[] conversionTasks = new Task[workbook.Worksheets.Count];
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            int sheetIndex = i; // capture loop variable for the lambda
            conversionTasks[i] = Task.Run(() =>
            {
                try
                {
                    // Get the worksheet
                    Worksheet sheet = workbook.Worksheets[sheetIndex];

                    // Determine the used range of the sheet
                    int maxRow = sheet.Cells.MaxDataRow;
                    int maxCol = sheet.Cells.MaxDataColumn;

                    // Create a range that covers the entire used area
                    Aspose.Cells.Range range = sheet.Cells.CreateRange(0, 0, maxRow + 1, maxCol + 1);

                    // Convert the range to HTML (uses Range.ToHtml rule)
                    byte[] htmlBytes = range.ToHtml(htmlOptions);

                    // Write the HTML bytes to a file named after the sheet
                    string htmlPath = Path.Combine(outputDir, $"Sheet{sheetIndex + 1}.html");
                    File.WriteAllBytes(htmlPath, htmlBytes);

                    Console.WriteLine($"Worksheet \"{sheet.Name}\" saved to {htmlPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing sheet index {sheetIndex}: {ex.Message}");
                }
            });
        }

        try
        {
            // Await completion of all conversion tasks
            await Task.WhenAll(conversionTasks);
            Console.WriteLine("All worksheets have been converted to HTML asynchronously.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during conversion: {ex.Message}");
        }
    }
}
