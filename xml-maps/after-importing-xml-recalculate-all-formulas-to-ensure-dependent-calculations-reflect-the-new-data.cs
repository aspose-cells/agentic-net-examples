using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Import XML data into the first worksheet starting at cell A1
        // Adjust the file path as needed; the sheet name must exist or will be created
        workbook.ImportXml("data.xml", "Sheet1", 0, 0);

        // Recalculate all formulas so that any dependent cells reflect the newly imported data
        workbook.CalculateFormula();

        // Refresh pivot tables and charts if they exist (optional but ensures full consistency)
        workbook.Worksheets.RefreshAll();

        // Save the updated workbook
        workbook.Save("Result.xlsx");
    }
}