using System;
using Aspose.Cells;
using System.Drawing;

class ReplaceThemeFontExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define the custom font family to be applied
            string customFontFamily = "Calibri";

            // Set the default style font to the custom font
            Style defaultStyle = workbook.DefaultStyle;
            defaultStyle.Font.Name = customFontFamily;
            defaultStyle.IsFontApplied = true; // Ensure font changes are applied for named style
            workbook.DefaultStyle = defaultStyle;

            // Add sample data to demonstrate the font change
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Original Font");
            sheet.Cells["A2"].PutValue("Will be changed to custom font");
            sheet.Cells["B1"].PutValue(123);
            sheet.Cells["B2"].PutValue(DateTime.Now);

            // Iterate through all worksheets and update each cell's style to use the custom font
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Use the used range to limit iteration to cells that contain data
                Aspose.Cells.Range usedRange = ws.Cells.MaxDisplayRange;
                int firstRow = usedRange.FirstRow;
                int firstColumn = usedRange.FirstColumn;
                int rowCount = usedRange.RowCount;
                int columnCount = usedRange.ColumnCount;

                for (int row = firstRow; row < firstRow + rowCount; row++)
                {
                    for (int col = firstColumn; col < firstColumn + columnCount; col++)
                    {
                        Cell cell = ws.Cells[row, col];
                        Style style = cell.GetStyle();

                        // Apply the custom font family
                        style.Font.Name = customFontFamily;
                        style.IsFontApplied = true; // Ensure the font change takes effect

                        cell.SetStyle(style);
                    }
                }
            }

            // Save the workbook with the updated theme font
            string outputPath = "WorkbookWithCustomThemeFont.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}