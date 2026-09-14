// Title: Hide zero values in rows after the 100th row using Aspose.Cells for .NET
// AI Prompts: Create C# code that defines a style with the format "0;-0;;@" and assigns it to every numeric cell from row 101 onward in each worksheet using Aspose.Cells. | Write a C# method that loops through all worksheets, finds numeric cells in rows greater than 100, applies a zero‑hiding style, and saves the workbook.
// Common Searches: Aspose.Cells hide zeros in rows after row 100 C# example | C# apply custom number format to suppress zero values starting at row 101 in Excel | How to set a style for numeric cells beyond a specific row using Aspose.Cells .NET | Iterate over used range and hide zero values in Excel with Aspose.Cells
// Tags: zero-suppression style Aspose.Cells | apply style to numeric cells beyond row 100 C# | hide zero values in Excel worksheet programmatically | iterate used range rows >100 Aspose.Cells

using Aspose.Cells;

// Loads a workbook, creates a style with the custom format "0;-0;;@" to hide zeros, applies this style to all numeric cells from row 101 to the last used row in each worksheet, and saves the updated file.
class HideZeroValuesBeyondRow100
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Define a style that hides zero values using a custom number format
        Style hideZeroStyle = workbook.CreateStyle();
        hideZeroStyle.Custom = "0;-0;;@"; // Positive;Negative;;Text (zero is hidden)

        // Apply the style to all cells in rows beyond the 100th row (row index 100 = 101st row)
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the last used row and column to limit the iteration
            int lastRow = sheet.Cells.MaxDataRow;
            int lastColumn = sheet.Cells.MaxDataColumn;

            // Iterate from row 101 (index 100) to the last used row
            for (int rowIndex = 100; rowIndex <= lastRow; rowIndex++)
            {
                // Apply the style to each cell in the current row
                for (int colIndex = 0; colIndex <= lastColumn; colIndex++)
                {
                    Cell cell = sheet.Cells[rowIndex, colIndex];
                    // Only apply to numeric cells; other types are unaffected by the format
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        cell.SetStyle(hideZeroStyle);
                    }
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
