using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class ProtectRowsExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the zero‑based start row (20th row) and the number of rows to protect (6 rows: 20‑25)
            int startRow = 19;          // Row 20 (zero‑based)
            int rowsCount = 6;          // Rows 20 to 25 inclusive
            int startColumn = 0;        // First column (A)

            // Create a temporary range covering the first column of the desired rows,
            // then expand it to the entire rows with the EntireRow property.
            AsposeRange rowsRange = sheet.Cells.CreateRange(startRow, startColumn, rowsCount, 1).EntireRow;

            // Add a protected range to the worksheet's AllowEditRanges collection.
            ProtectedRangeCollection allowRanges = sheet.AllowEditRanges;
            int protectedIndex = allowRanges.Add(
                "Rows20to25",
                rowsRange.FirstRow,
                rowsRange.FirstColumn,
                rowsRange.FirstRow + rowsRange.RowCount - 1,
                rowsRange.FirstColumn + rowsRange.ColumnCount - 1);

            ProtectedRange protectedRange = allowRanges[protectedIndex];
            protectedRange.Password = "MySecretPassword";

            // Protect the worksheet (required for the protected range to take effect)
            sheet.Protect(ProtectionType.All);

            // Save the workbook
            string outputPath = "ProtectedRows.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}