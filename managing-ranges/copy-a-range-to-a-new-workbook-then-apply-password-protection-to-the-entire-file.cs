// Title: Copy a specific cell range to a new Excel workbook and apply password protection with Aspose.Cells for .NET
// Common Searches: how to copy a range of cells to a new workbook and set a password using Aspose.Cells in C# | Aspose.Cells example for copying A1:C10 to another workbook and protecting the file | C# code to manually copy a cell range when CopyRange method is not available in Aspose.Cells | protect entire Excel workbook with a password after copying data using Aspose.Cells .NET | save a copied range to a new .xlsx file with password protection using Aspose.Cells
// Tags: copy cell range Aspose.Cells C# | manual range copy without CopyRange API | password protect workbook Aspose.Cells | protect entire Excel file Aspose.Cells | create new workbook from copied range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads source.xlsx, defines a cell area (A1:C10), creates an empty workbook, manually copies each cell's value and style to the new worksheet using nested loops, protects the whole destination workbook with a password via Protect(ProtectionType.All, password), and saves it as protected_copy.xlsx, handling missing files and exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string destPath = "protected_copy.xlsx";
            const string password = "MySecurePassword";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Load the source workbook containing the range to copy
            Workbook srcWorkbook = new Workbook(sourcePath);
            Worksheet srcWorksheet = srcWorkbook.Worksheets[0];

            // Define the source range (example: A1:C10)
            CellArea sourceArea = CellArea.CreateCellArea("A1", "C10");

            // Create a new, empty workbook for the copy
            Workbook destWorkbook = new Workbook();
            Worksheet destWorksheet = destWorkbook.Worksheets[0];

            // Copy the defined range manually (fallback when CopyRange API is unavailable)
            CopyRange(srcWorksheet, destWorksheet, sourceArea, 0, 0);

            // Apply password protection to the entire file
            destWorkbook.Protect(ProtectionType.All, password);

            // Save the new workbook with protection
            destWorkbook.Save(destPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    private static void CopyRange(Worksheet src, Worksheet dest, CellArea srcArea, int destStartRow, int destStartColumn)
    {
        int totalRows = srcArea.EndRow - srcArea.StartRow + 1;
        int totalColumns = srcArea.EndColumn - srcArea.StartColumn + 1;

        for (int i = 0; i < totalRows; i++)
        {
            for (int j = 0; j < totalColumns; j++)
            {
                int srcRow = srcArea.StartRow + i;
                int srcCol = srcArea.StartColumn + j;
                int destRow = destStartRow + i;
                int destCol = destStartColumn + j;

                Cell srcCell = src.Cells[srcRow, srcCol];
                Cell destCell = dest.Cells[destRow, destCol];

                // Copy value
                destCell.PutValue(srcCell.Value);

                // Copy style if needed
                destCell.SetStyle(srcCell.GetStyle());
            }
        }
    }
}
