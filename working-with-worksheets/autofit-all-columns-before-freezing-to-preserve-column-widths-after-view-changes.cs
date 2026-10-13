// Title: Auto-fit all columns in each worksheet then freeze panes to keep column widths using Aspose.Cells for .NET
// AI Prompts: Generate C# code that iterates through every worksheet, auto‑fits each used column, and then freezes the first row and first column with Aspose.Cells. | Create a method that determines the maximum column index per sheet, calls AutoFitColumn for every column up to that index, and applies FreezePanes so the column widths remain unchanged.
// Common Searches: C# Aspose.Cells auto fit all columns before freezing panes | How to preserve column widths when using FreezePanes in Aspose.Cells .NET | Iterate through worksheets to auto‑fit columns with Aspose.Cells | Find last used column index in a worksheet using Aspose.Cells | Freeze top row and left column after auto‑fitting columns Aspose.Cells
// Tags: auto-fit columns per worksheet Aspose.Cells | preserve column width with FreezePanes Aspose.Cells | worksheet iteration auto-fit Aspose.Cells | max column index detection Aspose.Cells | apply FreezePanes after column auto-fit Aspose.Cells

using Aspose.Cells;

// Loads an Excel workbook, auto‑fits every used column in all worksheets, freezes the first row and column on the first sheet, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Auto‑fit all columns in every worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the last used column (zero‑based)
            int lastCol = sheet.Cells.MaxColumn;
            // Auto‑fit each column up to the last used one
            for (int col = 0; col <= lastCol; col++)
            {
                sheet.AutoFitColumn(col);
            }
        }

        // Freeze panes after auto‑fit to preserve column widths
        // Example: freeze the first row and first column on the first worksheet
        Worksheet firstSheet = workbook.Worksheets[0];
        // FreezePanes(totalRows, totalColumns, row, column)
        // This freezes 1 row and 1 column; the scrollable area starts at cell B2 (row=1, column=1)
        firstSheet.FreezePanes(1, 1, 1, 1);

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
