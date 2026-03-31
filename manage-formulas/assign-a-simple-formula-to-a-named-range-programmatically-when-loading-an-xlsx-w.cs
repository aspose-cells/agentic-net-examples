using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeFormula
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Ensure the workbook has at least one worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Add a named range if it does not already exist
            int nameIndex = workbook.Worksheets.Names.Add("MyFormula");
            Name namedFormula = workbook.Worksheets.Names[nameIndex];

            // Assign a simple formula to the named range.
            // The formula must start with an equal sign.
            // Example: sum of cells A1 and A2 on Sheet1.
            namedFormula.RefersTo = "=Sheet1!$A$1+$A$2";

            // (Optional) Demonstrate usage: put the named formula in a cell.
            // The cell will display the result of the formula after calculation.
            sheet.Cells["B1"].Formula = "=MyFormula";

            // Calculate formulas so that the result is stored in the cell.
            workbook.CalculateFormula();

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            // Output verification
            Console.WriteLine("Named range 'MyFormula' set to: " + namedFormula.RefersTo);
            Console.WriteLine("Cell B1 value after calculation: " + sheet.Cells["B1"].Value);
        }
    }
}