using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data in columns C (index 2) to F (index 5)
        worksheet.Cells["C1"].PutValue("Short");
        worksheet.Cells["D1"].PutValue("A bit longer text");
        worksheet.Cells["E1"].PutValue("This is a considerably longer piece of text");
        worksheet.Cells["F1"].PutValue("Short again");

        // Auto‑fit columns C through F (zero‑based indices 2 to 5)
        worksheet.AutoFitColumns(2, 5);

        // Save the workbook
        workbook.Save("AutoFitColumns_C_F.xlsx");
    }
}