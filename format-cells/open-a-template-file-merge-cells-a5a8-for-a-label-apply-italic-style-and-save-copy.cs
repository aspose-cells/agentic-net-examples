using System;
using Aspose.Cells;

class MergeAndStyleExample
{
    static void Main()
    {
        // Path to the existing template workbook
        string templatePath = "Template.xlsx";

        // Path where the modified copy will be saved
        string outputPath = "Template_Modified.xlsx";

        // Load the template workbook
        Workbook workbook = new Workbook(templatePath);

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells A5:A8 (zero‑based indices: row 4, column 0, 4 rows, 1 column)
        worksheet.Cells.Merge(firstRow: 4, firstColumn: 0, totalRows: 4, totalColumns: 1);

        // Apply italic font style to the merged cell (upper‑left cell of the range)
        Style mergedCellStyle = worksheet.Cells[4, 0].GetStyle();
        mergedCellStyle.Font.IsItalic = true;
        worksheet.Cells[4, 0].SetStyle(mergedCellStyle);

        // Save the modified workbook as a new copy
        workbook.Save(outputPath);
    }
}