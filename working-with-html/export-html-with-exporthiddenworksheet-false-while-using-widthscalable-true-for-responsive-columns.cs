// Title: Export an Excel workbook to responsive HTML without hidden worksheets using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook and saves it as HTML while omitting hidden worksheets and enabling scalable column widths with Aspose.Cells. | Show how to set up Aspose.Cells HtmlSaveOptions to produce responsive HTML output and exclude hidden sheets in a .NET application.
// Common Searches: Aspose.Cells C# export to HTML exclude hidden worksheets | How to make column widths responsive when saving Excel as HTML with Aspose.Cells | HtmlSaveOptions WidthScalable true example in .NET | Generate HTML from an Excel file without hidden sheets using Aspose.Cells | Responsive HTML output from Excel workbook Aspose.Cells .NET
// Tags: HtmlSaveOptions ExportHiddenWorksheet false | WidthScalable true responsive columns Aspose.Cells | C# export Excel to HTML Aspose.Cells | omit hidden worksheets HTML export Aspose.Cells | responsive column sizing Aspose.Cells HTML

using Aspose.Cells;

// Loads an Excel workbook, configures HtmlSaveOptions to skip hidden worksheets and enable scalable column widths, then saves the file as responsive HTML using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Set up HTML export options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.ExportHiddenWorksheet = false; // Do not include hidden worksheets in the output
        htmlOptions.WidthScalable = true;          // Make column widths responsive for different screen sizes

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
