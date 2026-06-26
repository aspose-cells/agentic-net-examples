using System;
using System.Drawing;
using Aspose.Cells;
using CellsRange = Aspose.Cells.Range;   // Alias to avoid conflict with System.Range

namespace AsposeCellsThemeDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // -----------------------------------------------------------------
                // 1. Set a built‑in theme color (e.g., change Accent1 to a custom blue)
                // -----------------------------------------------------------------
                workbook.SetThemeColor(ThemeColorType.Accent1, Color.Blue);

                // -----------------------------------------------------------------
                // 2. Update the default style to use the new theme color.
                //    This makes the theme affect all cells that use the default style.
                // -----------------------------------------------------------------
                Style defaultStyle = workbook.DefaultStyle;
                defaultStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
                defaultStyle.Font.Name = "Calibri";
                defaultStyle.Font.Size = 11;
                workbook.DefaultStyle = defaultStyle; // assign back (lifecycle rule)

                // -----------------------------------------------------------------
                // 3. Apply the updated default style to existing cells.
                //    Iterate through each worksheet and its used range.
                // -----------------------------------------------------------------
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the used range of the worksheet.
                    CellsRange usedRange = sheet.Cells.MaxDisplayRange;
                    if (usedRange == null) continue;

                    int firstRow = usedRange.FirstRow;
                    int lastRow = firstRow + usedRange.RowCount - 1;
                    int firstCol = usedRange.FirstColumn;
                    int lastCol = firstCol + usedRange.ColumnCount - 1;

                    for (int row = firstRow; row <= lastRow; row++)
                    {
                        for (int col = firstCol; col <= lastCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];
                            Style cellStyle = cell.GetStyle();
                            // Apply the theme color from the default style.
                            cellStyle.Font.ThemeColor = defaultStyle.Font.ThemeColor;
                            cell.SetStyle(cellStyle);
                        }
                    }
                }

                // -----------------------------------------------------------------
                // 4. Save the workbook (lifecycle rule: save)
                // -----------------------------------------------------------------
                string outputPath = "WorkbookWithBuiltInTheme.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}