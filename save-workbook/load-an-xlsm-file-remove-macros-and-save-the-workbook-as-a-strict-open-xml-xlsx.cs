using System;
using Aspose.Cells;

namespace AsposeCellsMacroRemoval
{
    class Program
    {
        static void Main()
        {
            // Path to the macro‑enabled workbook (XLSM)
            string sourcePath = "input.xlsm";

            // Path for the macro‑free workbook (XLSX)
            string destinationPath = "output.xlsx";

            // Load the XLSM file
            Workbook workbook = new Workbook(sourcePath);

            // Remove all VBA/macros from the workbook
            workbook.RemoveMacro();

            // Save the workbook as a strict Open XML XLSX file
            workbook.Save(destinationPath, SaveFormat.Xlsx);

            Console.WriteLine($"Macros removed and file saved as: {destinationPath}");
        }
    }
}