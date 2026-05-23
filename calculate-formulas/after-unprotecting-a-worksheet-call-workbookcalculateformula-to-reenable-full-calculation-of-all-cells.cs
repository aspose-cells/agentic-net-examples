using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class UnprotectAndRecalculate
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (or any specific worksheet by index/name)
            Worksheet worksheet = workbook.Worksheets[0];

            // Unprotect the worksheet.
            // If the worksheet was protected with a password, provide it here.
            // If there is no password, you can call Unprotect() without arguments.
            worksheet.Unprotect("yourPassword"); // replace with actual password or use worksheet.Unprotect();

            // Re‑enable full calculation of all cells in the workbook.
            workbook.CalculateFormula();

            // Save the workbook after recalculation (replace with your desired output path)
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}