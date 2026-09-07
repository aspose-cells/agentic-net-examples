// Title: Export a workbook to HTML with a defined print area and suppress worksheet properties using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to save an Excel workbook as HTML, limiting the output to a defined print area and disabling worksheet property export. | Show how to set HtmlSaveOptions.ExportWorksheetProperties to false and configure Worksheet.PageSetup.PrintArea before calling Workbook.Save for HTML conversion.
// Common Searches: how to export a specific range to HTML with Aspose.Cells in C# | Aspose.Cells C# hide worksheet properties when saving as HTML | set print area for HTML output using Aspose.Cells .NET | save Excel workbook as HTML without worksheet metadata Aspose.Cells | Aspose.Cells HtmlSaveOptions ExportWorksheetProperties example
// Tags: Aspose.Cells HTML export with print area | HtmlSaveOptions ExportWorksheetProperties false | C# define worksheet print area Aspose.Cells | suppress worksheet properties Aspose.Cells HTML

using Aspose.Cells;

// The example creates a workbook, populates cells, defines the print area (A1:B2), configures HtmlSaveOptions to exclude worksheet properties, and saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Sample";

        // Populate some sample data
        sheet.Cells["A1"].PutValue("Header1");
        sheet.Cells["B1"].PutValue("Header2");
        sheet.Cells["A2"].PutValue(123);
        sheet.Cells["B2"].PutValue(456);

        // Define the print area (A1:B2)
        sheet.PageSetup.PrintArea = "A1:B2";

        // Set HTML save options to omit worksheet properties
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.ExportWorksheetProperties = false;

        // Export the workbook to HTML
        workbook.Save("output.html", htmlOptions);
    }
}
