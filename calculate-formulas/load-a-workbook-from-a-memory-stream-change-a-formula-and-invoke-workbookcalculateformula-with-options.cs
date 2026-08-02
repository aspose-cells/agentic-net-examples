using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a workbook with an initial formula and save it to a
        //    memory stream (simulating an existing file in memory).
        // ------------------------------------------------------------
        Workbook originalWorkbook = new Workbook();
        Worksheet originalSheet = originalWorkbook.Worksheets[0];
        originalSheet.Cells["A1"].PutValue(5);                 // A1 = 5
        originalSheet.Cells["B1"].Formula = "=A1*2";           // B1 = A1 * 2

        using (MemoryStream stream = new MemoryStream())
        {
            // Save the workbook into the stream.
            originalWorkbook.Save(stream, SaveFormat.Xlsx);
            // Reset the stream position so it can be read from the beginning.
            stream.Position = 0;

            // ------------------------------------------------------------
            // 2. Load the workbook from the memory stream.
            // ------------------------------------------------------------
            Workbook workbook = new Workbook(stream);
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // 3. Change an existing formula.
            //    - Add a new value in C1.
            //    - Update B1 to use the new cell.
            // ------------------------------------------------------------
            sheet.Cells["C1"].PutValue(10);                     // C1 = 10
            sheet.Cells["B1"].Formula = "=C1*3";                // B1 = C1 * 3

            // ------------------------------------------------------------
            // 4. Prepare calculation options.
            // ------------------------------------------------------------
            CalculationOptions calcOptions = new CalculationOptions
            {
                Recursive = true,      // calculate dependent cells across worksheets
                IgnoreError = false    // do not suppress calculation errors
            };

            // ------------------------------------------------------------
            // 5. Calculate all formulas using the specified options.
            // ------------------------------------------------------------
            workbook.CalculateFormula(calcOptions);

            // ------------------------------------------------------------
            // 6. Output the result of the changed formula.
            // ------------------------------------------------------------
            Console.WriteLine("Calculated B1 value: " + sheet.Cells["B1"].Value);

            // ------------------------------------------------------------
            // 7. (Optional) Save the modified workbook to a file.
            // ------------------------------------------------------------
            workbook.Save("ModifiedWorkbook.xlsx");
        }
    }
}