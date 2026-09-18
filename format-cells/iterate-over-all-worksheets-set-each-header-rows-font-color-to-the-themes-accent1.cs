// Title: Set header row font color to the workbook's Accent1 theme color across all worksheets using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells in C# to loop through every worksheet and change the font color of the first row to the workbook's ThemeColorType.Accent1. | Apply the workbook's Accent1 theme color to header cells, handling dynamic column ranges, and save the updated workbook.
// Common Searches: aspnet change first row font color to theme accent1 in each sheet Aspose.Cells | c# Aspose.Cells set header row color using workbook theme across multiple worksheets | how to apply workbook theme accent color to header cells in all worksheets with Aspose.Cells
// Tags: header row font styling Aspose.Cells | workbook theme accent1 color C# | loop through worksheets set cell style | dynamic column range cell formatting Aspose.Cells

// Load an existing workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("input.xlsx");

// Iterate over all worksheets in the workbook
foreach (Aspose.Cells.Worksheet sheet in workbook.Worksheets)
{
    // Assume the header row is the first row (index 0)
    // Determine the last used column in the sheet to limit the iteration
    int lastColumn = sheet.Cells.MaxDataColumn;

    // Loop through each cell in the header row
    for (int col = 0; col <= lastColumn; col++)
    {
        // Get the cell at row 0, current column
        Aspose.Cells.Cell cell = sheet.Cells[0, col];

        // Retrieve the current style of the cell
        Aspose.Cells.Style style = cell.GetStyle();

        // Set the font color to the workbook's theme Accent1 color
        style.Font.Color = workbook.GetThemeColor(Aspose.Cells.ThemeColorType.Accent1);

        // Apply the modified style back to the cell
        cell.SetStyle(style);
    }
}

// Save the modified workbook
workbook.Save("output.xlsx");
