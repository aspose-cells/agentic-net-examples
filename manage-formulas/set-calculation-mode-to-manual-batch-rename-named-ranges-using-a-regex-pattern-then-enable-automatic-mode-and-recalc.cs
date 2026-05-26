using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsRenameNamedRanges
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Sample data: create some named ranges to rename
                // -------------------------------------------------
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";

                // Populate sample cells
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["B1"].PutValue(30);
                sheet.Cells["B2"].PutValue(40);

                // Define named ranges (names must be valid Excel identifiers – no spaces)
                int idx1 = workbook.Worksheets.Names.Add("TotalSales");
                workbook.Worksheets.Names[idx1].RefersTo = "=Sheet1!$A$1:$A$2";

                int idx2 = workbook.Worksheets.Names.Add("QuarterlyReport");
                workbook.Worksheets.Names[idx2].RefersTo = "=Sheet1!$B$1:$B$2";

                int idx3 = workbook.Worksheets.Names.Add("YearlySummary");
                workbook.Worksheets.Names[idx3].RefersTo = "=Sheet1!$A$1:$B$2";

                // -------------------------------------------------
                // 1. Set calculation mode to Manual
                // -------------------------------------------------
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

                // -------------------------------------------------
                // 2. Batch rename named ranges using a regex pattern
                //    Example pattern: replace capital letters with underscores and prepend "New_"
                // -------------------------------------------------
                // This pattern inserts an underscore before each capital letter (except the first)
                Regex renamePattern = new Regex("(?<!^)([A-Z])");

                foreach (Name name in workbook.Worksheets.Names)
                {
                    string original = name.Text;                     // e.g., "TotalSales"
                    string withUnderscores = renamePattern.Replace(original, "_$1"); // "Total_Sales"
                    string newName = "New_" + withUnderscores;       // "New_Total_Sales"
                    name.Text = newName;
                }

                // -------------------------------------------------
                // 3. Switch calculation mode back to Automatic
                // -------------------------------------------------
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

                // -------------------------------------------------
                // 4. Recalculate all formulas in the workbook
                // -------------------------------------------------
                workbook.CalculateFormula();

                // -------------------------------------------------
                // Save the workbook (lifecycle rule: save)
                // -------------------------------------------------
                workbook.Save("RenamedNamedRanges.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}