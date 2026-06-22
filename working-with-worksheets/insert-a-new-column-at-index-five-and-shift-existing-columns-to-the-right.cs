using Aspose.Cells;

class InsertColumnAtIndexFive
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Fill some sample data in the first few rows and columns
        for (int row = 0; row < 5; row++)
        {
            for (int col = 0; col < 8; col++)
            {
                cells[row, col].PutValue($"R{row}C{col}");
            }
        }

        // Insert a new column at index 5 (0‑based). Existing columns 5 and beyond shift right.
        cells.InsertColumn(5);

        // Add a header to the newly inserted column (optional)
        cells[0, 5].PutValue("New Column");

        // Save the workbook
        workbook.Save("InsertColumnAtIndexFive.xlsx");
    }
}