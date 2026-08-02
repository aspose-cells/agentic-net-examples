using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate the source range E5:G10 with sample formulas
        // Rows 5-10 correspond to indices 4-9, columns E-G correspond to indices 4-6
        for (int row = 4; row <= 9; row++)
        {
            for (int col = 4; col <= 6; col++)
            {
                cells[row, col].Formula = $"=SUM(A{row + 1}:B{row + 1})";
            }
        }

        // Define the source CellArea for E5:G10
        CellArea sourceArea = new CellArea
        {
            StartRow = 4,      // Row 5 (zero‑based)
            StartColumn = 4,   // Column E (zero‑based)
            EndRow = 9,        // Row 10
            EndColumn = 6      // Column G
        };

        // Destination start cell J5 -> row index 4, column index 9
        int destRow = 4;      // Row 5
        int destColumn = 9;   // Column J

        // Move the range while preserving formulas
        cells.MoveRange(sourceArea, destRow, destColumn);

        // Save the workbook
        workbook.Save("MovedRange.xlsx");
    }
}