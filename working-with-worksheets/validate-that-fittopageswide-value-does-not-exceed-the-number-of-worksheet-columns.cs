using System;
using Aspose.Cells;

namespace AsposeCellsFitToPagesWideValidation
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data to define used columns
            // This creates 5 used columns (A to E)
            for (int col = 0; col < 5; col++)
            {
                worksheet.Cells[0, col].PutValue($"Header {col + 1}");
                worksheet.Cells[1, col].PutValue($"Data {col + 1}");
            }

            // Set the FitToPagesWide value (example: 3 pages wide)
            worksheet.PageSetup.FitToPagesWide = 3;

            // Validation: ensure FitToPagesWide does not exceed the number of used columns
            // Cells.MaxColumn returns the zero‑based index of the last used column.
            int usedColumnCount = worksheet.Cells.MaxColumn + 1; // convert to count

            if (worksheet.PageSetup.FitToPagesWide > usedColumnCount)
            {
                // If invalid, adjust to the maximum allowed value or raise an error.
                // Here we choose to reset it to the column count.
                Console.WriteLine($"FitToPagesWide ({worksheet.PageSetup.FitToPagesWide}) exceeds used columns ({usedColumnCount}). Adjusting value.");
                worksheet.PageSetup.FitToPagesWide = usedColumnCount;
            }
            else
            {
                Console.WriteLine($"FitToPagesWide ({worksheet.PageSetup.FitToPagesWide}) is within the used column range ({usedColumnCount}).");
            }

            // Save the workbook (lifecycle: save)
            workbook.Save("FitToPagesWideValidated.xlsx");
        }
    }
}