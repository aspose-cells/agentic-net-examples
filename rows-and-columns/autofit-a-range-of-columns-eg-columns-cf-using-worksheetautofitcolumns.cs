using System;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – auto‑fit columns C‑F
class Program
{
    static void Main()
    {
        // Create a new workbook (empty)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Sample data to demonstrate auto‑fit (optional)
        worksheet.Cells["C1"].PutValue("Short");
        worksheet.Cells["D1"].PutValue("Medium length text");
        worksheet.Cells["E1"].PutValue("A very very long piece of text that requires more width");
        worksheet.Cells["F1"].PutValue(12345);

        // Auto‑fit columns C (index 2) through F (index 5)
        worksheet.AutoFitColumns(2, 5);

        // Save the workbook
        workbook.Save("AutoFitColumns_C_F.xlsx");
    }
}