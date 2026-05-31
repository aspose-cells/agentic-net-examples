using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ProtectWorksheetAllowFormulaEdit
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data: some cells with values, some with formulas
            cells["A1"].PutValue(10);
            cells["A2"].Formula = "=A1*2";
            cells["B1"].PutValue(5);
            cells["B2"].Formula = "=B1+3";
            cells["C1"].PutValue("Text");
            cells["C2"].Formula = "=CONCATENATE(A1,B1)";

            // Unlock only the cells that contain formulas
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.IsFormula)
                    {
                        // Get the current style, set IsLocked = false, and apply it back
                        Style style = cell.GetStyle();
                        style.IsLocked = false;
                        cell.SetStyle(style);
                    }
                }
            }

            // Protect the worksheet (all protection types) with a password
            // Locked cells cannot be edited, but unlocked formula cells can be edited
            sheet.Protect(ProtectionType.All, "securePassword", null);

            // Save the workbook
            string outputPath = "ProtectedAllowFormulaEdit.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
    }
}