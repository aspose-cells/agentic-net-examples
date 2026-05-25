using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // For SaveFormat enum

namespace RecalculateFormulasBeforeConversion
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Convert the workbook to PDF (you can change the format as needed)
            workbook.Save("output.pdf", SaveFormat.Pdf);

            Console.WriteLine("Formulas recalculated and workbook converted to PDF successfully.");
        }
    }
}