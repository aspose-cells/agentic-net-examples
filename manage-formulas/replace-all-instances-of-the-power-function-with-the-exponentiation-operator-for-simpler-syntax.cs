using System;
using Aspose.Cells;
using System.Text.RegularExpressions;

class ReplacePowerFunction
{
    static void Main()
    {
        // Load the workbook
        Workbook wb = new Workbook("input.xlsx");

        // Regex to transform POWER(x,y) into (x)^(y)
        Regex powerRegex = new Regex(@"POWER\(([^,]+),([^\)]+)\)", RegexOptions.IgnoreCase);

        // Iterate through all worksheets and cells
        foreach (Worksheet sheet in wb.Worksheets)
        {
            Cells cells = sheet.Cells;
            foreach (Cell cell in cells)
            {
                if (cell.IsFormula && cell.Formula.IndexOf("POWER", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    string originalFormula = cell.Formula;
                    string transformedFormula = powerRegex.Replace(originalFormula, "($1)^($2)");
                    cell.Formula = transformedFormula;
                }
            }
        }

        // Recalculate all formulas after replacement
        wb.CalculateFormula();

        // Save the modified workbook
        wb.Save("output.xlsx");
    }
}