using System;
using Aspose.Cells;

class CleanPdfReport
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Remove all pivot tables from every worksheet in the workbook
        workbook.Worksheets.ClearPivottables();

        // Recalculate formulas to ensure the PDF reflects the latest data
        workbook.CalculateFormula();

        // Save the cleaned workbook as a PDF file
        workbook.Save("CleanReport.pdf", SaveFormat.Pdf);
    }
}