using System;
using System.IO;
using Aspose.Cells;

class ApplyCustomNumberFormat
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate cells with formulas (some produce negative results)
            cells["A1"].Formula = "=5-10";          // -5
            cells["A2"].Formula = "=20-5";          // 15
            cells["B1"].Formula = "=SUM(-3,2)";     // -1
            cells["B2"].Formula = "=SUM(4,6)";      // 10

            // Calculate all formulas so that cell values are up‑to‑date
            workbook.CalculateFormula();

            // Custom number format that shows negative numbers in red
            string customNumberFormat = "_-\"$\"* #,##0.00_ ;[Red]-\"$\"* #,##0.00_ ;_-\"$\"* \"-\"??_ ;_(@_)";

            // Create a style that contains only the custom number format
            Style negativeStyle = workbook.CreateStyle();
            negativeStyle.Custom = customNumberFormat;

            // Use StyleFlag to apply only the number format part of the style
            StyleFlag numberFormatFlag = new StyleFlag { NumberFormat = true };

            // Iterate through all cells, find those with formulas that evaluate to a negative number,
            // and apply the custom number format using the style flag
            foreach (Cell cell in cells)
            {
                if (cell.IsFormula && cell.Type == CellValueType.IsNumeric && cell.DoubleValue < 0)
                {
                    // Create a one‑cell range and apply the style with the flag
                    Aspose.Cells.Range cellRange = sheet.Cells.CreateRange(cell.Row, cell.Column, 1, 1);
                    cellRange.ApplyStyle(negativeStyle, numberFormatFlag);
                }
            }

            // Save the workbook (ensure the directory exists)
            string outputPath = "NegativeFormulaFormat.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}