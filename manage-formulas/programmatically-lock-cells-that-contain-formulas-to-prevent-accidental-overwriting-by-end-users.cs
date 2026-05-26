using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class LockFormulaCellsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data
                cells["A1"].PutValue(10);
                cells["A2"].PutValue(20);
                cells["A3"].PutValue(30);

                // Add formulas that we want to protect
                cells["B1"].Formula = "=A1*2";
                cells["B2"].Formula = "=A2*2";
                cells["B3"].Formula = "=A3*2";

                // Lock cells containing formulas
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell.IsFormula)
                        {
                            Style style = cell.GetStyle();
                            style.IsLocked = true;
                            cell.SetStyle(style);
                        }
                    }
                }

                // Protect the worksheet (oldPassword parameter required – pass empty string)
                worksheet.Protect(ProtectionType.All, "SecurePassword123", string.Empty);

                // Save the workbook
                string outputPath = "LockedFormulaCells.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            LockFormulaCellsDemo.Run();
        }
    }
}