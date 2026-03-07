using System;
using Aspose.Cells;

namespace AsposeCellsPrintTitleClearDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (or any specific worksheet)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the PageSetup object for the worksheet
            PageSetup pageSetup = worksheet.PageSetup;

            // Clear the print title rows and columns by setting them to an empty string
            pageSetup.PrintTitleRows = string.Empty;      // Clears rows to repeat on each printed page
            pageSetup.PrintTitleColumns = string.Empty;   // Clears columns to repeat on each printed page

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}