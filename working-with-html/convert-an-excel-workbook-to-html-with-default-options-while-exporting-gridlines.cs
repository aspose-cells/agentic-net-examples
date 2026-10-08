// Title: Convert an Excel workbook to HTML with default options and export gridlines using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, sets HtmlSaveOptions.ExportGridLines to true, and saves it as an HTML file with Aspose.Cells. | Show how to use Aspose.Cells HtmlSaveOptions to export gridlines while converting a workbook to HTML with default settings. | Provide a minimal Aspose.Cells example that converts a workbook to HTML, preserving the original gridlines.
// Common Searches: Aspose.Cells C# export Excel to HTML with gridlines | how to enable gridline export when saving workbook as HTML using Aspose.Cells | default HtmlSaveOptions settings for converting .xlsx to .html in .NET | sample code to convert Excel file to HTML preserving gridlines Aspose.Cells | HtmlSaveOptions ExportGridLines true example C#
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines | C# convert Excel to HTML with gridlines | Aspose.Cells default HTML conversion options | export Excel gridlines to HTML .NET | save workbook as HTML using Aspose.Cells

using System;
using Aspose.Cells;

// // Loads 'input.xlsx', enables gridline export via HtmlSaveOptions, and saves the workbook as 'output.html' using Aspose.Cells.
class ExcelToHtmlConverter
{
    static void Main()
    {
        // Load the Excel workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to export gridlines
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportGridLines = true; // Enable gridline export

        // Save the workbook as an HTML file with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
