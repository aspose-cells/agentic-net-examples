using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // default workbook with one worksheet

        // Get the first worksheet and its cells collection
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Determine how many rows to unhide.
        // MaxDataRow returns the last row that contains data (zero‑based).
        // Add 1 to get the count, and ensure at least one row is processed.
        int totalRows = cells.MaxDataRow + 1;
        if (totalRows == 0) totalRows = 1;

        // Unhide all rows from the first row (index 0) to the determined count.
        // Height = -1 means auto‑fit the row height after unhiding.
        cells.UnhideRows(0, totalRows, -1);

        // Enable formula display so formulas are shown instead of their results.
        worksheet.ShowFormulas = true;

        // Save the modified workbook to a new file.
        workbook.Save("UnhiddenFormulas.xlsx", SaveFormat.Xlsx);
    }
}