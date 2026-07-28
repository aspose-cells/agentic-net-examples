// Title: Set Workbook Theme to a Monospaced Font (Consolas) for All Cells with Aspose.Cells for .NET
// Description: Creates a workbook, changes the default style to the monospaced Consolas font with a Minor scheme, updates every existing cell to use the same font while keeping values intact, and saves the file.
// Keywords: Aspose.Cells | C# | default font | monospaced font | Consolas | workbook theme | font scheme Minor | apply font to all cells | Excel report formatting | code snippet display
// Common Searches: Aspose.Cells change default workbook font C# | apply monospaced font to all cells Aspose.Cells | set font scheme to Minor in Aspose.Cells workbook | update workbook theme font programmatically | C# code to use Consolas for Excel cells
// Developer Intent: Apply a monospaced font to the workbook theme and propagate it to every cell.
// Use Cases: Render code snippets in Excel reports with a consistent monospaced typeface for better readability. | Ensure newly added cells automatically inherit the Consolas default style after the change. | Retrofit an existing workbook so all previously formatted cells adopt the new monospaced font without altering their data.
// AI Prompts: Generate C# Aspose.Cells code that sets the workbook default font to Consolas and updates every cell to use that font. | Show how to change the font scheme to Minor and apply a monospaced font across an entire workbook using Aspose.Cells. | Explain step‑by‑step how to modify a workbook’s theme font in Aspose.Cells while preserving existing cell values.

using System;
using Aspose.Cells;

namespace AsposeCellsThemeFontUpdate
{
    // Creates a workbook, changes the default style to the monospaced Consolas font with a Minor scheme, updates every existing cell to use the same font while keeping values intact, and saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // Define a monospaced font name (e.g., Consolas)
                const string monospacedFont = "Consolas";

                // Update the default style so that any new cells inherit the monospaced font
                Style defaultStyle = workbook.DefaultStyle;
                defaultStyle.Font.Name = monospacedFont;
                defaultStyle.Font.SchemeType = FontSchemeType.Minor;
                workbook.DefaultStyle = defaultStyle;

                // Apply the updated style to all existing cells in all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the used range of the worksheet
                    Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                    if (usedRange == null) continue;

                    // Iterate through each cell in the used range
                    for (int row = usedRange.FirstRow; row <= usedRange.FirstRow + usedRange.RowCount - 1; row++)
                    {
                        for (int col = usedRange.FirstColumn; col <= usedRange.FirstColumn + usedRange.ColumnCount - 1; col++)
                        {
                            Cell cell = sheet.Cells[row, col];
                            // Preserve existing cell value, only modify the style
                            Style cellStyle = cell.GetStyle();
                            cellStyle.Font.Name = monospacedFont;
                            cellStyle.Font.SchemeType = FontSchemeType.Minor;
                            cell.SetStyle(cellStyle);
                        }
                    }
                }

                // Save the workbook (lifecycle: save)
                workbook.Save("WorkbookWithMonospacedTheme.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
