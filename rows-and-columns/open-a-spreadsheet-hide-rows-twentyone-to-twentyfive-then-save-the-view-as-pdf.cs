using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing spreadsheet
        Workbook workbook = new Workbook("input.xlsx"); // replace with your file path

        // Get the first worksheet (or specify another index/name as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Hide rows 21 through 25 (zero‑based index starts at 20, hide 5 rows)
        worksheet.Cells.HideRows(20, 5);

        // Save the workbook view as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}