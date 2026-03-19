using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetManipulation
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            string sourcePath = "SourceWorkbook.xlsx";
            Workbook workbook = new Workbook(sourcePath);

            // -------------------------------------------------
            // 1. Duplicate a worksheet by its name
            // -------------------------------------------------
            // Use the first worksheet's name to ensure it exists
            string sheetToCopy = workbook.Worksheets[0].Name;

            // AddCopy returns the index of the newly created copy
            int copiedIndex = workbook.Worksheets.AddCopy(sheetToCopy);
            Worksheet copiedSheet = workbook.Worksheets[copiedIndex];

            // Give the copied sheet a distinct name
            copiedSheet.Name = sheetToCopy + "_Copy";

            // -------------------------------------------------
            // 2. Relocate the copied worksheet to a specific position
            // -------------------------------------------------
            // Move the copied sheet to index 1 (second position in the workbook)
            copiedSheet.MoveTo(1);

            // -------------------------------------------------
            // 3. Optional: Swap two worksheets by index
            // -------------------------------------------------
            // For demonstration, swap the first sheet with the copied sheet
            // (after MoveTo, the copied sheet is at index 1)
            workbook.Worksheets.SwapSheet(0, 1);

            // -------------------------------------------------
            // 4. Save the modified workbook
            // -------------------------------------------------
            string outputPath = "ModifiedWorkbook.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine("Worksheet duplication and relocation completed successfully.");
        }
    }
}