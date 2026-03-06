using System;
using Aspose.Cells;

class ConvertTextToNumberDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the cells collection of the first worksheet
        Cells cells = workbook.Worksheets[0].Cells;

        // Insert values that are stored as text
        cells["A1"].PutValue("123");          // integer as text
        cells["A2"].PutValue("45.67");        // decimal as text
        cells["A3"].PutValue("2021-06-20");   // date as text
        cells["A4"].PutValue("NotANumber");   // non‑numeric text

        // Convert all convertible string values to their numeric/date equivalents
        cells.ConvertStringToNumericValue();

        // Export the workbook to XLSX format
        workbook.Save("ConvertedNumbers.xlsx", SaveFormat.Xlsx);
    }
}