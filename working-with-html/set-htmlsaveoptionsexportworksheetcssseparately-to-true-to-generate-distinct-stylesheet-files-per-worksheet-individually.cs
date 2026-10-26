// Title: Create a separate CSS stylesheet for each worksheet when converting an Excel file to HTML with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export each worksheet to its own CSS file when saving as HTML | how to generate individual stylesheet per sheet using HtmlSaveOptions in .NET | set ExportWorksheetCSSSeparately to true for Excel to HTML conversion Aspose.Cells | separate CSS files per worksheet Aspose.Cells HTMLSaveOptions example | C# convert XLSX to HTML with distinct CSS per worksheet using Aspose
// Tags: HtmlSaveOptions separate stylesheet per sheet | Aspose.Cells per-worksheet CSS export | Excel workbook HTML conversion with distinct CSS files | C# generate individual CSS files from worksheets | Aspose.Cells HTML export stylesheet separation

using Aspose.Cells;

// The sample loads an Excel workbook, enables HtmlSaveOptions.ExportWorksheetCSSSeparately, and saves the file as HTML, causing Aspose.Cells to produce a separate CSS file for each worksheet.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

        // Set to export a separate CSS file for each worksheet
        htmlOptions.ExportWorksheetCSSSeparately = true;

        // Save the workbook to HTML; separate CSS files will be generated per worksheet
        workbook.Save("output.html", htmlOptions);
    }
}
