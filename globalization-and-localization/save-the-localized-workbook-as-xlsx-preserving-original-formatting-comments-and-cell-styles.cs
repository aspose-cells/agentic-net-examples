using System;
using Aspose.Cells;

namespace AsposeCellsSaveLocalizedWorkbook
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook (could be any localized Excel file)
            string sourcePath = "input.xlsx";

            // Path where the workbook will be saved preserving all formatting, comments, and styles
            string destinationPath = "output.xlsx";

            // Load the existing workbook
            Workbook workbook = new Workbook(sourcePath);

            // Save the workbook as XLSX. The Save method with SaveFormat preserves
            // original formatting, comments, and cell styles automatically.
            workbook.Save(destinationPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{destinationPath}'.");
        }
    }
}