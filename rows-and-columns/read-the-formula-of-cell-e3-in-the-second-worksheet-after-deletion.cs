using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class ReadFormulaAfterDeletion
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Delete the third row (index 2) from the first worksheet.
            // The second parameter ensures that references in other worksheets are updated.
            Worksheet firstSheet = workbook.Worksheets[0];
            firstSheet.Cells.DeleteRow(2, true);

            // Access the second worksheet (index 1) and read the formula in cell E3.
            Worksheet secondSheet = workbook.Worksheets[1];
            string formula = secondSheet.Cells["E3"].Formula;

            // Output the retrieved formula.
            Console.WriteLine($"Formula in Sheet2!E3 after deletion: {formula}");

            // Save the modified workbook (optional).
            workbook.Save("output.xlsx");
        }
    }
}