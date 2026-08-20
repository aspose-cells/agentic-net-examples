// Title: Export Aspose.Cells Workbook to Separate HTML Files While Preserving Sheet Order (C#)
// Description: Loads an Excel workbook, configures HtmlSaveOptions with a custom IFilePathProvider that returns "{sheetName}.html", and saves each worksheet as an individual HTML file in the same sequence as the source workbook.
// Keywords: Aspose.Cells | C# | .NET | HtmlSaveOptions | IFilePathProvider | export workbook to HTML | preserve worksheet order | separate HTML per sheet | Excel to HTML conversion | custom file path provider
// Common Searches: Aspose.Cells export workbook to HTML preserve sheet order | C# HtmlSaveOptions custom IFilePathProvider example | Save each Excel worksheet as separate HTML file Aspose | How to keep original sheet sequence when converting Excel to HTML .NET | Aspose.Cells generate one HTML file per worksheet
// Developer Intent: Create HTML output for every worksheet of an Excel file without altering the original sheet sequence by using a custom IFilePathProvider.
// Use Cases: Web‑based reporting where each Excel sheet must appear as its own HTML page in the same order as the workbook. | Documentation systems that require individual HTML files named after sheet titles while maintaining the workbook's logical flow. | Integration of Excel data into a portal that loads separate HTML files per sheet and relies on the original worksheet ordering.
// AI Prompts: Show a C# snippet that exports an Aspose.Cells workbook to HTML using a custom IFilePathProvider that keeps the sheet order unchanged. | Explain how to modify PreserveOrderFilePathProvider to write HTML files into a subfolder while still preserving the original worksheet sequence. | Demonstrate configuring HtmlSaveOptions to embed images as base64 strings when using the custom file path provider.

using System;
using Aspose.Cells;

// Loads an Excel workbook, configures HtmlSaveOptions with a custom IFilePathProvider that returns "{sheetName}.html", and saves each worksheet as an individual HTML file in the same sequence as the source workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Set a custom IFilePathProvider that returns a file name based on the worksheet name.
        // This keeps the original worksheet order because the provider does not modify the mapping.
        saveOptions.FilePathProvider = new PreserveOrderFilePathProvider();

        // Save the workbook to HTML. Each worksheet will be saved as a separate HTML file
        // (e.g., Sheet1.html, Sheet2.html, ...) while preserving the original order.
        workbook.Save("output.html", saveOptions);
    }

    // Custom implementation of IFilePathProvider
    private class PreserveOrderFilePathProvider : IFilePathProvider
    {
        // Returns the full file name for a given worksheet name.
        // No reordering or custom naming logic is applied.
        public string GetFullName(string sheetName)
        {
            return $"{sheetName}.html";
        }
    }
}
