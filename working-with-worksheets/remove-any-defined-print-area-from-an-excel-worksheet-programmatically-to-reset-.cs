using System;
using Aspose.Cells;

namespace AsposeCellsPrintAreaReset
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets and clear the defined print area
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Setting PrintArea to an empty string removes any previously defined print area
                sheet.PageSetup.PrintArea = string.Empty;
            }

            // Save the workbook with the print area cleared (replace with desired output path)
            workbook.Save("output.xlsx");
        }
    }
}