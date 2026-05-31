using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Number of columns the header should span (e.g., A to E)
        int totalColumns = 5; // columns are zero‑based, so this spans columns 0‑4

        // Merge the first row across the specified columns
        // Parameters: firstRow, firstColumn, totalRows, totalColumns
        cells.Merge(0, 0, 1, totalColumns);

        // Set the header text in the merged cell
        cells[0, 0].PutValue("Report Header");

        // Apply a style to center the text and make it bold
        Style headerStyle = workbook.CreateStyle();
        headerStyle.HorizontalAlignment = TextAlignmentType.Center;
        headerStyle.VerticalAlignment = TextAlignmentType.Center;
        headerStyle.Font.IsBold = true;
        cells[0, 0].SetStyle(headerStyle);

        // Save the workbook
        workbook.Save("HeaderMerged.xlsx");
    }
}