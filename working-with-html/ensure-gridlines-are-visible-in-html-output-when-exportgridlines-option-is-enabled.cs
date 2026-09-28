// Title: How to export an Excel workbook to HTML with visible gridlines using Aspose.Cells for .NET
// AI Prompts: Write C# code that saves a workbook as an HTML file and enables gridline rendering with Aspose.Cells. | Show the exact steps to set HtmlSaveOptions.ExportGridLines to true for HTML export in Aspose.Cells. | Adapt a basic Aspose.Cells HTML export example so that cell borders appear in the generated HTML output.
// Common Searches: Aspose.Cells C# export worksheet to HTML with gridlines enabled | HtmlSaveOptions ExportGridLines property not showing gridlines in HTML output | display Excel gridlines in HTML using Aspose.Cells .NET | how to make gridlines visible when saving Excel as HTML with Aspose.Cells
// Tags: Aspose.Cells HTML export gridlines | HtmlSaveOptions ExportGridLines property | C# export workbook to HTML with borders | display Excel gridlines in HTML using Aspose.Cells | Aspose.Cells HTMLSaveOptions gridline visibility

using Aspose.Cells;

// The example creates a workbook, adds sample data, configures HtmlSaveOptions.ExportGridLines = true, and saves the file as ExportedWithGridLines.html, resulting in an HTML representation that shows the worksheet's gridlines.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Populate some sample data to illustrate gridlines
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Header1");
        sheet.Cells["B1"].PutValue("Header2");
        sheet.Cells["A2"].PutValue(123);
        sheet.Cells["B2"].PutValue(456);

        // Configure HTML export options to include gridlines
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportGridLines = true; // Enable gridline visibility in the HTML output

        // Save the workbook as an HTML file with the specified options
        workbook.Save("ExportedWithGridLines.html", htmlOptions);
    }
}
