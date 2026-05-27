using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (uses the standard creation pattern)
        Workbook workbook = new Workbook();

        // Add extra worksheets for demonstration purposes
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");

        // Apply A4 paper size to every worksheet using a foreach loop
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
        }

        // Save the workbook (uses the standard save pattern)
        workbook.Save("A4PaperSizeWorkbook.xlsx", SaveFormat.Xlsx);
    }
}