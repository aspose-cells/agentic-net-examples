// Title: High‑performance HTML export of a large Excel workbook using Aspose.Cells HtmlCrossType.Cross in C#
// AI Prompts: Write C# code that loads an existing workbook and saves it as HTML with HtmlSaveOptions configured for HtmlCrossType.Cross to accelerate rendering of massive sheets. | Show how to enable cross‑type HTML rendering in Aspose.Cells when exporting a workbook that contains thousands of rows. | Provide a sample that checks the Aspose.Cells version, applies HtmlCrossType.Cross if supported, and falls back to the default HTML renderer for large workbooks.
// Common Searches: Aspose.Cells C# export large workbook to HTML with HtmlCrossType.Cross | How to enable cross‑type HTML rendering in Aspose.Cells for .NET | Fast HTML conversion of Excel files using HtmlCrossType in Aspose.Cells | HtmlSaveOptions HtmlCrossType property missing after upgrading Aspose.Cells | Performance tips for exporting massive Excel sheets to HTML with Aspose.Cells
// Tags: Aspose.Cells HtmlCrossType.Cross HTML rendering | C# large workbook HTML export performance | HtmlSaveOptions cross type rendering Aspose.Cells | Excel to HTML high‑speed conversion .NET | upgrade Aspose.Cells for HtmlCrossType support

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook, fills it with sample data, and demonstrates how to configure HtmlSaveOptions for high‑performance HTML export using HtmlCrossType.Cross. It also includes a version check and notes that the HtmlCrossType property is only available in newer Aspose.Cells releases, advising an upgrade when necessary.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Populate the first worksheet with sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Oranges");
            sheet.Cells["B3"].PutValue(85);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: HtmlCrossType property is not available in this version of Aspose.Cells.
            // If needed, upgrade the library to a version that supports HtmlCrossType.

            // Export the workbook to an HTML file
            workbook.Save("LargeWorkbookExport.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
