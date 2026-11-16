// Title: Generate a separate CSS file for each worksheet when saving to HTML with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# save workbook to HTML with separate CSS per worksheet | How to create individual stylesheet files for each Excel sheet using HtmlSaveOptions | Set ExportWorksheetCSSSeparately to true in Aspose.Cells C# example | Generate multiple CSS files when converting Excel to HTML with Aspose.Cells | C# Aspose.Cells HTML export distinct stylesheet per worksheet
// Tags: HtmlSaveOptions ExportWorksheetCSSSeparately | Aspose.Cells per-worksheet CSS generation | C# save Excel as HTML with separate stylesheets | Excel to HTML conversion distinct CSS files | multiple stylesheet files per worksheet Aspose.Cells

using Aspose.Cells;

// The example creates a workbook, adds data, enables HtmlSaveOptions.ExportWorksheetCSSSeparately, and saves the workbook as HTML, causing Aspose.Cells to produce an individual CSS file for each worksheet.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Configure HTML save options to export CSS separately for each worksheet
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        saveOptions.ExportWorksheetCSSSeparately = true;

        // Save the workbook as HTML; distinct stylesheet files will be generated per worksheet
        workbook.Save("output.html", saveOptions);
    }
}
