// Title: Replace Workbook Theme Font with a Custom Font Family and Update All Cell Styles – Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, sets the default style font to a custom family (e.g., Calibri) with a major scheme, then walks through every worksheet and used cell to apply the same font to each cell's style before saving the file as CustomThemeFont.xlsx.
// Keywords: Aspose.Cells | C# | change theme font | custom font family | update cell styles | default workbook font | Excel automation | font scheme major | global font change
// Common Searches: Aspose.Cells change theme font C# | set default font for entire workbook Aspose.Cells | apply custom font to all cells Aspose.Cells .NET | replace theme font in Excel using Aspose | global font update Aspose.Cells example
// Developer Intent: Replace the workbook’s theme font with a specific custom font family and ensure that every existing cell style in all worksheets reflects the new font.
// Use Cases: Generate corporate Excel reports that automatically use the company’s standard font without manual formatting. | Retrofitting legacy workbooks to match a new branding guideline by updating the theme and all cell styles in one operation. | Automating multilingual Excel exports where a single, universally supported font must be enforced across all sheets.
// AI Prompts: Show C# code using Aspose.Cells to change the workbook theme font to 'Arial' and propagate the change to all existing cell styles. | Provide an example that iterates through every worksheet and cell in an Aspose.Cells workbook, setting each cell's font to a custom family while preserving other style attributes. | Explain how to update the default style and apply a major font scheme to all cells in an Excel file with Aspose.Cells for .NET.

using Aspose.Cells;
using Aspose.Cells;
using System;

// Creates a new Workbook, sets the default style font to a custom family (e.g., Calibri) with a major scheme, then walks through every worksheet and used cell to apply the same font to each cell's style before saving the file as CustomThemeFont.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data
        worksheet.Cells["A1"].PutValue("Sample Text 1");
        worksheet.Cells["A2"].PutValue("Sample Text 2");

        // Define the custom font family to replace the theme font
        string customFontFamily = "Calibri";

        // Update the default style of the workbook
        Style defaultStyle = workbook.DefaultStyle;
        defaultStyle.Font.Name = customFontFamily;
        defaultStyle.Font.SchemeType = FontSchemeType.Major; // apply to major scheme
        workbook.DefaultStyle = defaultStyle;

        // Iterate through all worksheets and cells to update existing styles
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the used range
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    if (cell != null)
                    {
                        Style style = cell.GetStyle();
                        style.Font.Name = customFontFamily;
                        style.Font.SchemeType = FontSchemeType.Major;
                        cell.SetStyle(style);
                    }
                }
            }
        }

        // Save the workbook (lifecycle rule: save)
        workbook.Save("CustomThemeFont.xlsx");
    }
}
