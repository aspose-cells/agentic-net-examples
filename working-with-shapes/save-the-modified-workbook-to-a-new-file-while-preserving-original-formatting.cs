using System;
using Aspose.Cells;

class SaveWorkbookExample
{
    static void Main()
    {
        // Load an existing workbook (preserves all original formatting)
        Workbook workbook = new Workbook("input.xlsx");

        // Example modification: change the value of cell A1 in the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Modified");

        // Save the modified workbook to a new file while keeping original formatting
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}