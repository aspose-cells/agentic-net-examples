// Title: How to export an Excel workbook to HTML with visible gridlines using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# export workbook to HTML with gridlines visible | Enable gridlines in HTML output using HtmlSaveOptions ExportGridLines | C# Aspose.Cells HTML export gridlines not showing | Save Excel as HTML with borders using Aspose.Cells | Configure HtmlSaveOptions to include cell borders in generated HTML
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines property | C# generate HTML from Excel with cell borders | HTML gridline rendering using Aspose.Cells | Excel to HTML conversion preserving gridlines | Aspose.Cells HTML export configuration

using Aspose.Cells;

// The example creates a Workbook, fills cells with data, sets HtmlSaveOptions.ExportGridLines to true, and saves the file as HTML, ensuring that gridlines appear in the generated HTML output.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue("Header1");
        sheet.Cells["B1"].PutValue("Header2");
        sheet.Cells["A2"].PutValue(10);
        sheet.Cells["B2"].PutValue(20);
        sheet.Cells["A3"].PutValue(30);
        sheet.Cells["B3"].PutValue(40);

        // Configure HTML save options to export gridlines
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportGridLines = true; // Enable gridlines in the HTML output

        // Save the workbook as an HTML file with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
