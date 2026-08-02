using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Author: Aspose.Cells .NET example – AutoFitRows for a specific range of rows
    class AutoFitRowsRangeDemo
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: creation)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data that will affect row heights
            worksheet.Cells["A2"].PutValue("This is a long text that should cause row 2 to expand.");
            worksheet.Cells["B2"].PutValue("Additional content in column B.");
            worksheet.Cells["A3"].PutValue("Short");
            worksheet.Cells["B3"].PutValue("Another short text.");
            worksheet.Cells["A4"].PutValue("A very very long piece of text that will make row 4 taller than default.");
            worksheet.Cells["B4"].PutValue("More text to increase height.");

            // AutoFit rows 2 through 4 (zero‑based indices 1 to 3)
            worksheet.AutoFitRows(startRow: 1, endRow: 3);

            // Save the workbook (lifecycle rule: saving)
            workbook.Save("AutoFitRowsRangeDemo.xlsx");
        }
    }
}