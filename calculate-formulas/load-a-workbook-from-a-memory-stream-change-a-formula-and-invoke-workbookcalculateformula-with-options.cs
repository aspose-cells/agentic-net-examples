using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing Excel file into a byte array (replace with your source)
        byte[] excelBytes = File.ReadAllBytes("input.xlsx");

        // Create a memory stream from the byte array
        using (MemoryStream inputStream = new MemoryStream(excelBytes))
        {
            // Load the workbook from the memory stream
            Workbook workbook = new Workbook(inputStream);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Change the formula of a specific cell (e.g., B2)
            // Original formula is replaced with a new one that multiplies A1 by 10
            sheet.Cells["B2"].Formula = "=A1*10";

            // Prepare calculation options (e.g., ignore errors during calculation)
            CalculationOptions calcOptions = new CalculationOptions
            {
                IgnoreError = true
            };

            // Recalculate all formulas in the workbook using the specified options
            workbook.CalculateFormula(calcOptions);

            // Save the updated workbook to a new memory stream
            using (MemoryStream outputStream = new MemoryStream())
            {
                workbook.Save(outputStream, SaveFormat.Xlsx);

                // Optionally write the result to a physical file for verification
                File.WriteAllBytes("output.xlsx", outputStream.ToArray());
            }
        }
    }
}