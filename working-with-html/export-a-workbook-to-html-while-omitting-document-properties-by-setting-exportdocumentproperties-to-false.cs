// Title: Convert an Excel workbook to HTML without embedding document properties using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file and saves it as HTML with HtmlSaveOptions.ExportDocumentProperties set to false using Aspose.Cells. | Show how to configure Aspose.Cells HtmlSaveOptions to exclude workbook metadata when exporting to HTML. | Generate a minimal example that converts a workbook to HTML while omitting document properties in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# export Excel to HTML without document properties | How to disable ExportDocumentProperties in HtmlSaveOptions when saving as HTML | Save workbook as HTML without metadata using Aspose.Cells .NET | Remove workbook properties from HTML output Aspose.Cells example | HtmlSaveOptions ExportDocumentProperties false C#
// Tags: Aspose.Cells HtmlSaveOptions ExportDocumentProperties false | C# export Excel to HTML without metadata | disable workbook properties in HTML conversion Aspose.Cells | HTML export options Aspose.Cells .NET | convert XLSX to HTML without document properties

using Aspose.Cells;

// Loads 'input.xlsx', sets HtmlSaveOptions.ExportDocumentProperties = false, and saves the workbook as 'output.html' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the source workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to omit document properties
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportDocumentProperties = false;

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
