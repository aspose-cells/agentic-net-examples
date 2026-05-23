using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsThemeBorderDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // ------------------------------------------------------------
                // Create a style that uses the theme's Accent2 color for all borders
                // ------------------------------------------------------------
                Style borderStyle = workbook.CreateStyle();

                // Set line style for visibility (optional)
                borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                // Apply the theme's Accent2 color to each border via ThemeColor property.
                ThemeColor accent2Theme = new ThemeColor(ThemeColorType.Accent2, 0);
                borderStyle.Borders[BorderType.TopBorder].ThemeColor = accent2Theme;
                borderStyle.Borders[BorderType.BottomBorder].ThemeColor = accent2Theme;
                borderStyle.Borders[BorderType.LeftBorder].ThemeColor = accent2Theme;
                borderStyle.Borders[BorderType.RightBorder].ThemeColor = accent2Theme;

                // ------------------------------------------------------------
                // Prepare a StyleFlag to indicate that only border settings are applied
                // ------------------------------------------------------------
                StyleFlag flag = new StyleFlag { Borders = true };

                // ------------------------------------------------------------
                // Apply the style to every worksheet's used range (covers all tables)
                // ------------------------------------------------------------
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Determine the used range; if the sheet is empty, apply to a single cell (A1)
                    int maxRow = sheet.Cells.MaxDataRow >= 0 ? sheet.Cells.MaxDataRow + 1 : 1;
                    int maxCol = sheet.Cells.MaxDataColumn >= 0 ? sheet.Cells.MaxDataColumn + 1 : 1;

                    // Fully qualify Range to avoid ambiguity with System.Range
                    Aspose.Cells.Range usedRange = sheet.Cells.CreateRange(0, 0, maxRow, maxCol);
                    usedRange.ApplyStyle(borderStyle, flag);
                }

                // ------------------------------------------------------------
                // Save the workbook
                // ------------------------------------------------------------
                string outputPath = "WorkbookWithThemeAccent2Borders.xlsx";

                // Ensure the directory exists (handle possible null from GetDirectoryName)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}