using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source rows with values and formulas
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["B1"].Formula = "=A1*2";   // Formula that should be ignored
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["B2"].Formula = "=A2*2";   // Another formula

            // Define the source range covering the rows to copy (A1:B2)
            AsposeRange sourceRange = sheet.Cells.CreateRange(0, 0, 2, 2); // rows 0-1, columns 0-1

            // Define the destination range where rows will be pasted (starting at row 5)
            AsposeRange destinationRange = sheet.Cells.CreateRange(4, 0, 2, 2); // rows 4-5, columns 0-1

            // Configure PasteOptions to copy only values (ignore formulas and formatting)
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.Values,   // Copy only cell values
                SkipBlanks = true,
                OnlyVisibleCells = false,
                Transpose = false,
                IgnoreLinksToOriginalFile = true
            };

            // Perform the copy operation using the specified paste options
            destinationRange.Copy(sourceRange, pasteOptions);

            // Define output file path
            string outputPath = "CopyValuesOnly.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}