using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Import XML data into the first worksheet starting at cell A1 (row 0, column 0)
        // Adjust the file path and sheet name as needed
        workbook.ImportXml("data.xml", "Sheet1", 0, 0);

        // Recalculate all formulas in the workbook so that dependent cells reflect the imported data
        workbook.CalculateFormula();

        // Save the updated workbook
        workbook.Save("Result.xlsx");
    }
}