using System;
using Aspose.Cells;

class UngroupColumnsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data in columns A to D
        cells["A1"].PutValue("Short");
        cells["B1"].PutValue("Medium length text");
        cells["C1"].PutValue("Very very long text that needs column width adjustment");
        cells["D1"].PutValue("Another example");

        // Set initial column widths to a consistent value (e.g., 15 characters)
        for (int col = 0; col <= 3; col++)
        {
            cells.Columns[col].Width = 15;
        }

        // Group columns B and C (indexes 1 to 2) and hide them as an example modification
        cells.GroupColumns(1, 2, true);

        // Perform some modification while columns are grouped
        cells["A2"].PutValue("New data while columns are grouped");

        // Store current column widths before ungrouping
        double[] storedWidths = new double[4];
        for (int col = 0; col <= 3; col++)
        {
            storedWidths[col] = cells.Columns[col].Width;
        }

        // Ungroup columns B and C (indexes 1 to 2)
        cells.UngroupColumns(1, 2);

        // Restore the stored widths to keep column widths consistent across the worksheet
        for (int col = 0; col <= 3; col++)
        {
            cells.Columns[col].Width = storedWidths[col];
        }

        // Save the workbook
        workbook.Save("UngroupColumnsConsistentWidths.xlsx");
    }
}