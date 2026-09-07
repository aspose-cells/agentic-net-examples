// Title: Convert an Excel workbook to HTML and include hidden worksheets using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file and saves it as .html with all worksheets, including hidden ones, using Aspose.Cells. | Show how to enable HtmlSaveOptions.ExportHiddenWorksheet in Aspose.Cells to export hidden sheets during HTML conversion. | Provide a minimal .NET example that converts a workbook to HTML while preserving hidden worksheet visibility.
// Common Searches: Aspose.Cells export hidden worksheets to HTML in C# | C# convert Excel file to HTML with hidden sheets included | HtmlSaveOptions ExportHiddenWorksheet property usage example | How to save an Excel workbook as HTML and keep hidden worksheets using Aspose.Cells | Convert .xlsx to .html with all sheets visible using Aspose.Cells .NET
// Tags: Aspose.Cells HTML hidden worksheet export | HtmlSaveOptions ExportHiddenWorksheet usage | C# Excel to HTML conversion Aspose | include hidden sheets in HTML output | Aspose.Cells save workbook as HTML

using Aspose.Cells;

// // Loads 'input.xlsx', sets HtmlSaveOptions.ExportHiddenWorksheet to true, and saves the workbook as 'output.html' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the Excel workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to include hidden worksheets
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportHiddenWorksheet = true; // export hidden sheets

        // Save the workbook as an HTML file with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
