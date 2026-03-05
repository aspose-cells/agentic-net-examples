using System;
using Aspose.Cells;

class ConvertNumbersStoredAsText
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet's cells
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate cells with numbers stored as text and a non‑numeric string
        cells["A1"].PutValue("123");          // numeric text
        cells["A2"].PutValue("456.78");       // numeric text with decimal
        cells["A3"].PutValue("NotANumber");   // will remain as string
        cells["A4"].PutValue("2021-06-20");   // date text

        // Convert all convertible string values to their numeric/date equivalents
        cells.ConvertStringToNumericValue();

        // Export the workbook to XLSX format
        workbook.Save("ConvertedNumbers.xlsx", SaveFormat.Xlsx);
    }
}