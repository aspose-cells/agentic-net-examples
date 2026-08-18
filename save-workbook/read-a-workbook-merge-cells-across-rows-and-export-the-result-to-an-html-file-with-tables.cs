// Title: C# – Merge Cells and Export Active Worksheet as HTML Table with Aspose.Cells
// Description: Load an Excel file, merge a range (e.g., A1:C3) on the first worksheet, set a header value, configure HtmlSaveOptions to force empty TD merging and export only the active sheet, then save the result as a compact HTML table.
// Keywords: Aspose.Cells C# merge cells | export worksheet to HTML Aspose | HtmlSaveOptions MergeEmptyTdType | save Excel as HTML table | active worksheet only HTML export
// Common Searches: Aspose.Cells merge range and save as HTML | C# export single worksheet to HTML table | how to reduce HTML size with Aspose.Cells | merge empty TD elements Aspose HTML export | C# Aspose.Cells HTMLSaveOptions examples
// Developer Intent: Combine selected cells in an Excel workbook and generate a single‑sheet HTML table.
// Use Cases: Create a spanning header for web‑displayed reports. | Produce lightweight HTML by collapsing consecutive empty cells. | Embed only the primary worksheet in a web page without extra sheets.
// AI Prompts: Generate C# code that merges A1:C3, adds a header, and saves the workbook as an HTML table using Aspose.Cells with forced empty‑TD merging. | Explain the impact of HtmlSaveOptions.MergeEmptyTdType.MergeForcely on the output HTML and when to apply it. | Adapt the sample to export each worksheet of a workbook as separate HTML files.

using System;
using Aspose.Cells;

// Load an Excel file, merge a range (e.g., A1:C3) on the first worksheet, set a header value, configure HtmlSaveOptions to force empty TD merging and export only the active sheet, then save the result as a compact HTML table.
class Program
{
    static void Main()
    {
        // Load the source workbook from a file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet and its cells collection
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge a range of cells across rows (e.g., A1:C3)
        // Parameters: firstRow (0‑based), firstColumn (0‑based), totalRows (1‑based), totalColumns (1‑based)
        cells.Merge(0, 0, 3, 3);

        // Optionally place a value in the merged cell (top‑left corner of the range)
        cells[0, 0].PutValue("Merged Header");

        // Set up HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        // Reduce HTML size by merging contiguous empty TD elements
        htmlOptions.MergeEmptyTdType = MergeEmptyTdType.MergeForcely;
        // Export only the active worksheet (as a table)
        htmlOptions.ExportActiveWorksheetOnly = true;

        // Save the workbook as an HTML file
        string outputPath = "output.html";
        workbook.Save(outputPath, htmlOptions);
    }
}
