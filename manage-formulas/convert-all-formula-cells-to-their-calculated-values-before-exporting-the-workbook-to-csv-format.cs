using System;
using Aspose.Cells;

namespace AsposeCellsFormulaToCsv
{
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Replace formulas with their calculated values for each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.Cells.RemoveFormulas();
            }

            // Save the workbook as CSV (formulas are now values)
            workbook.Save("output.csv", SaveFormat.Csv);
        }
    }
}