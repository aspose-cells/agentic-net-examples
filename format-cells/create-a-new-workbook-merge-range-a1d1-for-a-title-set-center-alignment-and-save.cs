using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells A1:D1 (row 0, column 0, 1 row, 4 columns)
        worksheet.Cells.Merge(0, 0, 1, 4);

        // Set the title text in the merged cell
        worksheet.Cells[0, 0].PutValue("Title");

        // Retrieve the style of the merged cell
        Style style = worksheet.Cells[0, 0].GetStyle();

        // Set horizontal alignment to center
        style.HorizontalAlignment = TextAlignmentType.Center;

        // Apply the updated style back to the cell
        worksheet.Cells[0, 0].SetStyle(style);

        // Save the workbook
        workbook.Save("MergedTitle.xlsx");
    }
}