// Title: How to export hidden worksheets to HTML with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx workbook, sets HtmlSaveOptions.ExportHiddenWorksheet to true, and saves it as an .html file using Aspose.Cells. | Show how to configure Aspose.Cells HtmlSaveOptions so that hidden Excel sheets are included when converting a workbook to HTML in C#. | Update an existing Aspose.Cells HTML export routine to ensure hidden worksheets appear in the generated HTML output.
// Common Searches: Aspose.Cells C# export hidden worksheet to HTML example | How to include hidden Excel sheets in HTML output using Aspose.Cells .NET | HtmlSaveOptions ExportHiddenWorksheet property usage in C# | Save workbook as HTML with hidden sheets visible Aspose.Cells | Convert Excel to HTML preserving hidden worksheets Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportHiddenWorksheet | C# export hidden worksheets to HTML | Aspose.Cells include hidden sheets in HTML conversion | HTML export of hidden Excel worksheets using Aspose.Cells | Aspose.Cells workbook to HTML with hidden sheets

using Aspose.Cells;

// The sample loads an Excel workbook, enables HtmlSaveOptions.ExportHiddenWorksheet to include hidden sheets, and saves the workbook as an HTML file, ensuring hidden worksheets are rendered in the output.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to include hidden worksheets
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportHiddenWorksheet = true; // Ensure hidden sheets are exported

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
