// Title: Export an Excel workbook to native‑style HTML using Aspose.Cells HtmlCrossType.MSExport in C#
// Common Searches: Aspose.Cells C# export workbook to HTML with HtmlCrossType.MSExport | How to generate Excel‑compatible HTML using HtmlSaveOptions in .NET | Save an Excel workbook as native style HTML preserving formatting with Aspose.Cells | Example code for HtmlCrossType MSExport HTML export in C# Aspose.Cells | Convert worksheet to HTML retaining Excel styles using Aspose.Cells API
// Tags: HtmlCrossType.MSExport Aspose.Cells HTML export | C# HtmlSaveOptions preserving Excel formatting | export workbook to native style HTML .NET | Aspose.Cells generate Excel compatible HTML | save worksheet as HTML with original styles

using System;
using System.IO;
using Aspose.Cells;

// The example creates a Workbook, adds sample data, configures HtmlSaveOptions with HtmlCrossType.MSExport to mimic Excel's native HTML output, and saves the workbook as an HTML file, demonstrating how to preserve Excel formatting when exporting to HTML using Aspose.Cells for .NET.
class ExportWorkbookToHtml
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(85);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // Optional: export only the active worksheet
            // htmlOptions.ExportActiveWorksheetOnly = true;
            // Optional: embed images as Base64
            // htmlOptions.ExportImagesAsBase64 = true;

            // Define output file path
            string outputPath = "ExportedWorkbook.html";

            // Save the workbook as an HTML file using the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully exported to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
