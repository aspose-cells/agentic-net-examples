using Aspose.Cells;

class MergeHeaderExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge cells A2:D2 (zero‑based row 1, column 0, 1 row, 4 columns)
        cells.Merge(1, 0, 1, 4);

        // Set the header text in the merged cell
        cells[1, 0].Value = "Header";

        // Create a style that centers the text horizontally and vertically
        Style style = cells[1, 0].GetStyle();
        style.HorizontalAlignment = TextAlignmentType.Center;
        style.VerticalAlignment = TextAlignmentType.Center;

        // Apply the style to the merged cell
        cells[1, 0].SetStyle(style);

        // Save the workbook
        workbook.Save("MergedHeader.xlsx");
    }
}