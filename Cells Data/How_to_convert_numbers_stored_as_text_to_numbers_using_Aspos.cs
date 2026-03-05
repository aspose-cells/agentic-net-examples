using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet's cells collection
        Cells cells = workbook.Worksheets[0].Cells;

        // Sample data: numbers stored as text and a non‑numeric string
        cells[0, 0].PutValue("123");      // will be converted to numeric
        cells[0, 1].PutValue("45.67");   // will be converted to numeric
        cells[0, 2].PutValue("ABC");     // remains as text

        // Convert all string values that can be interpreted as numbers
        cells.ConvertStringToNumericValue();

        // Export the workbook to XLSX format
        workbook.Save("ConvertedNumbers.xlsx", SaveFormat.Xlsx);
    }
}