// Title: Measure memory usage during large Excel‑to‑HTML conversion with CSS (Aspose.Cells for .NET)
// Description: Creates a 5,000‑row × 50‑column workbook, applies MemoryPreference, records process private memory before and after saving to HTML with Css enabled (HtmlSaveOptions), and outputs the memory delta in megabytes.
// Keywords: Aspose.Cells memory profiling | HTML conversion memory usage | .NET workbook to HTML | Css enabled HtmlSaveOptions | large Excel workbook performance | MemoryPreference Aspose.Cells | process PrivateMemorySize64 | benchmark Aspose.Cells HTML export
// Common Searches: Aspose.Cells memory usage when saving to HTML | How to track .NET memory during Excel to HTML conversion | HtmlSaveOptions CSS impact on memory consumption | Measure peak memory for large workbook HTML export | Reduce memory footprint Aspose.Cells HTML conversion
// Developer Intent: Quantify the memory consumed by Aspose.Cells when converting a massive workbook to HTML with CSS enabled.
// Use Cases: Benchmark memory impact of different HtmlSaveOptions settings. | Validate that MemoryPreference lowers peak memory for huge worksheets. | Compare memory usage with CSS enabled versus disabled during HTML export.
// AI Prompts: Generate C# code that logs GC collections, working set, and private memory before and after Aspose.Cells workbook.Save to HTML with CSS options. | Suggest additional techniques (streaming, partial saves, workbook splitting) to further reduce memory consumption for large Excel‑to‑HTML conversions. | Create a unit test that asserts the memory increase stays below a defined threshold when converting a 5,000 × 50 workbook to HTML with CSS enabled.

using System;
using System.Diagnostics;
using Aspose.Cells;

// Creates a 5,000‑row × 50‑column workbook, applies MemoryPreference, records process private memory before and after saving to HTML with Css enabled (HtmlSaveOptions), and outputs the memory delta in megabytes.
class MeasureMemoryDuringHtmlConversion
{
    static void Main()
    {
        // Create a large workbook with many rows and columns
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate the worksheet with sample data (e.g., 5000 rows x 50 columns)
        int rows = 5000;
        int cols = 50;
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                cells[r, c].PutValue($"R{r + 1}C{c + 1}");
            }
        }

        // Optional: set memory preference to reduce memory footprint during processing
        workbook.Settings.MemorySetting = MemorySetting.MemoryPreference;

        // Capture memory usage before conversion
        Process proc = Process.GetCurrentProcess();
        long memoryBefore = proc.PrivateMemorySize64;

        // Configure HTML save options with CSS enabled (default behavior)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Ensure CSS is used (inline styles only when DisableCss is true)
            DisableCss = false,
            // Export worksheet CSS separately to keep CSS files (optional)
            ExportWorksheetCSSSeparately = true,
            // Enable additional CSS custom properties for better performance (optional)
            EnableCssCustomProperties = true
        };

        // Convert the workbook to HTML
        string htmlPath = "LargeWorkbook.html";
        workbook.Save(htmlPath, htmlOptions);

        // Capture memory usage after conversion
        long memoryAfter = proc.PrivateMemorySize64;

        // Output memory usage information
        Console.WriteLine($"Memory before conversion: {memoryBefore / (1024 * 1024)} MB");
        Console.WriteLine($"Memory after conversion : {memoryAfter / (1024 * 1024)} MB");
        Console.WriteLine($"Memory increase         : {(memoryAfter - memoryBefore) / (1024 * 1024)} MB");
        Console.WriteLine($"HTML file saved to: {htmlPath}");
    }
}
