// Title: How to include hidden worksheets when converting an Excel workbook to HTML with Aspose.Cells for .NET
// AI Prompts: Create a C# program that loads an .xlsx file, enables the hidden‑worksheet export flag, and saves the workbook as a single HTML document using Aspose.Cells. | Show how to set the Aspose.Cells HtmlSaveOptions to include hidden sheets when converting an Excel workbook to HTML in .NET. | Write C# code that iterates over hidden worksheets in a workbook and saves each one to its own HTML file with Aspose.Cells.
// Common Searches: How to use Aspose.Cells to export hidden worksheets when saving to HTML in C# | Example of HtmlSaveOptions ExportHiddenWorksheet property in .NET | Convert an Excel file to HTML while preserving hidden sheets using Aspose.Cells | C# code to include hidden sheets in HTML output with Aspose.Cells | Aspose.Cells HTML conversion hidden worksheets tutorial
// Tags: Aspose.Cells hidden worksheet HTML export | C# Aspose.Cells generate HTML with hidden sheets | HTML conversion include hidden worksheets .NET | save Excel as HTML preserving hidden sheets | Aspose.Cells HtmlSaveOptions hidden sheet option

using Aspose.Cells;

// Loads an Excel workbook, sets HtmlSaveOptions.ExportHiddenWorksheet to true to include hidden sheets, and saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to include hidden worksheets
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportHiddenWorksheet = true; // Ensures hidden sheets are exported

        // Save the workbook as HTML with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
