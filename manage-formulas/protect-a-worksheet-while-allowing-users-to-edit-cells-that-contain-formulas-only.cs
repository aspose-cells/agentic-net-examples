// Title: Protect an Excel worksheet but keep only formula cells editable using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# protect worksheet while allowing edits only in cells that contain formulas | unlock cells with formulas before applying worksheet protection in Aspose.Cells .NET | C# code to protect Excel sheet and make formula cells editable using Aspose.Cells | how to set IsLocked false for formula cells and protect sheet with password in Aspose.Cells | selective worksheet protection based on formula presence Aspose.Cells C#
// Tags: selective cell unlocking for formula cells Aspose.Cells | worksheet protection with password Aspose.Cells C# | cell style IsLocked manipulation Aspose.Cells | iterate used range to detect formulas Aspose.Cells | apply protection type All Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// The example loads an existing workbook, scans the used range to find cells that contain formulas, unlocks those cells by setting their Style.IsLocked property to false, then protects the entire worksheet with a password so only the unlocked formula cells remain editable, and finally saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range
            AsposeRange usedRange = sheet.Cells.MaxDisplayRange;
            int startRow = usedRange.FirstRow;
            int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
            int startCol = usedRange.FirstColumn;
            int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

            // Unlock cells that contain formulas
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    if (cell.IsFormula)
                    {
                        Style style = cell.GetStyle();
                        style.IsLocked = false; // Unlock the cell
                        cell.SetStyle(style);
                    }
                }
            }

            // Protect the worksheet; only unlocked cells remain editable
            sheet.Protect(ProtectionType.All, "myPassword", string.Empty);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
