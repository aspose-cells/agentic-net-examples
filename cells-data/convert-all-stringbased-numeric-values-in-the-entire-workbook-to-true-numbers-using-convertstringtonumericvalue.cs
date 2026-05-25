using System;
using Aspose.Cells;

namespace ConvertStringNumbers
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Convert all string values that can be interpreted as numbers
                // (including dates) to true numeric values in the current sheet.
                sheet.Cells.ConvertStringToNumericValue();
            }

            // Save the modified workbook (replace with desired output path)
            workbook.Save("output.xlsx");
        }
    }
}