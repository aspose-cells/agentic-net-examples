using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet and its cells collection
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge a range that spans multiple rows (e.g., A1:C3)
        // Parameters: firstRow (0‑based), firstColumn (0‑based), totalRows (1‑based), totalColumns (1‑based)
        cells.Merge(0, 0, 3, 3);

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        // Reduce HTML size by merging contiguous empty TD elements
        htmlOptions.MergeEmptyTdType = MergeEmptyTdType.MergeForcely;
        // Export only the table part of the worksheet
        htmlOptions.ExportDataOptions = HtmlExportDataOptions.Table;

        // Save the workbook as an HTML file with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}