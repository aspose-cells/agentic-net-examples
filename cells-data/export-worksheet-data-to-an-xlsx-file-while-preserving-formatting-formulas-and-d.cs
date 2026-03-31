using System;
using Aspose.Cells;

class ExportWorksheet
{
    static void Main()
    {
        // Load the source workbook (preserves all data, formatting, and formulas)
        Workbook workbook = new Workbook("source.xlsx");

        // Ensure all formulas are calculated before saving (optional but keeps data integrity)
        workbook.CalculateFormula();

        // Save the workbook as XLSX, preserving formatting, formulas, and data
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}