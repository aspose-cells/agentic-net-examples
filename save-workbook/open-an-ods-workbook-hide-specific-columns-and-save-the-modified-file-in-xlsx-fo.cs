using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source ODS workbook
        Workbook workbook = new Workbook("input.ods");

        // Access the first worksheet (you can change the index as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Hide specific columns.
        // Example: hide columns B and C (zero‑based indexes 1 and 2)
        // HideColumns(startColumnIndex, numberOfColumns)
        worksheet.Cells.HideColumns(1, 2);

        // Save the modified workbook in XLSX format
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}