using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsVLookupToXLookup
{
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all cells that contain formulas
                foreach (Cell cell in cells)
                {
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula;

                        // Pattern matches VLOOKUP with the fourth argument set to FALSE (exact match)
                        // Example: VLOOKUP(A2,$B$2:$D$10,3,FALSE)
                        string pattern = @"VLOOKUP\(\s*([^,]+)\s*,\s*([^,]+)\s*,\s*([^,]+)\s*,\s*FALSE\s*\)";
                        if (Regex.IsMatch(formula, pattern, RegexOptions.IgnoreCase))
                        {
                            // Convert to XLOOKUP:
                            // XLOOKUP(lookup_value, INDEX(table_array,0,1), INDEX(table_array,0,col_index_num))
                            string replacement = "XLOOKUP($1,INDEX($2,0,1),INDEX($2,0,$3))";

                            string newFormula = Regex.Replace(formula, pattern, replacement, RegexOptions.IgnoreCase);

                            // Set the new formula back to the cell
                            cell.Formula = newFormula;
                        }
                    }
                }
            }

            // Recalculate all formulas after conversion
            workbook.CalculateFormula();

            // Save the modified workbook (replace with desired output path)
            workbook.Save("output.xlsx");
        }
    }
}