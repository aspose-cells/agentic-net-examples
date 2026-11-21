// Title: Export an Excel workbook to HTML with default settings while retaining conditional formatting using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export Excel to HTML preserving conditional formatting | How to save .xlsx as .html with default options in Aspose.Cells | C# convert workbook to HTML using Aspose.Cells without custom HtmlSaveOptions | Default HtmlSaveOptions in Aspose.Cells include conditional formatting | Save Excel file as HTML in .NET while keeping cell styles
// Tags: Aspose.Cells HTML export default options | Workbook.Save HtmlSaveOptions C# | Export Excel conditional formatting to HTML Aspose.Cells | C# convert XLSX to HTML Aspose.Cells | SaveFormat.Html Aspose.Cells .NET

using Aspose.Cells;

// The code loads an Excel file (input.xlsx) into a Workbook object and saves it as HTML (output.html) using SaveFormat.Html, which applies the default HtmlSaveOptions that automatically preserve conditional formatting in the generated HTML.
class Program
{
    static void Main()
    {
        // Load the Excel workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Save the workbook as HTML using default options.
        // Default HtmlSaveOptions export conditional formatting, so no extra configuration is needed.
        workbook.Save("output.html", SaveFormat.Html);
    }
}
