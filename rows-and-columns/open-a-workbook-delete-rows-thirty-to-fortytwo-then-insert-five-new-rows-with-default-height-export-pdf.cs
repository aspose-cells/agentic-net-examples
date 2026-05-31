using System;
using Aspose.Cells;

namespace AsposeCellsRowManipulation
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Delete rows 30 to 42 (inclusive). Row indices are zero‑based.
            // Start index = 30, total rows = 42 - 30 + 1 = 13
            cells.DeleteRows(30, 13);

            // Insert five new rows at the same position (row index 30).
            // These rows will have the default height.
            cells.InsertRows(30, 5);

            // Save the modified workbook as PDF
            workbook.Save("output.pdf", SaveFormat.Pdf);
        }
    }
}