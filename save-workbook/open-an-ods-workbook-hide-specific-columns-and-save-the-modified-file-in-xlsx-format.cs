using System;
using Aspose.Cells;

namespace AsposeCellsOdsToXlsx
{
    class Program
    {
        static void Main()
        {
            // Path to the source ODS file
            string sourcePath = "input.ods";

            // Load the ODS workbook
            Workbook workbook = new Workbook(sourcePath);

            // Access the first worksheet (you can modify this index as needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Hide a single column (e.g., column B, zero‑based index 1)
            cells.HideColumn(1);

            // Hide multiple columns (e.g., columns D, E, F – start at index 3, hide 3 columns)
            cells.HideColumns(3, 3);

            // Save the modified workbook as XLSX
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook converted from ODS to XLSX and saved to '{outputPath}'.");
        }
    }
}