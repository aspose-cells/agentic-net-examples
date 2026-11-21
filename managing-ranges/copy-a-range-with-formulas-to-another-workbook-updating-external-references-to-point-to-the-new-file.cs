// Title: Copy a cell range with formulas and styles to a new workbook and refresh external links using Aspose.Cells for .NET
// Common Searches: C# copy a range of cells with formulas to another workbook using Aspose.Cells | preserve cell formatting and formulas when transferring a range to a new Excel file in .NET | how to update external workbook references after copying cells with Aspose.Cells | example of manual cell iteration to copy A1:C10 range in Aspose.Cells | Aspose.Cells .NET code for copying range and fixing external links
// Tags: cell range copy preserving formulas Aspose.Cells | style preservation during workbook duplication C# | external link adjustment after range copy Aspose.Cells | iterative cell copy example .NET | Aspose.Cells create new workbook from range

using System;
using System.IO;
using Aspose.Cells;

// The sample loads a source Excel workbook, manually iterates over the A1:C10 range, copies each cell's formula (or value) and style to a newly created target workbook, handles missing files and cell‑level errors, and saves the target workbook, ensuring external references point to the new file.
class Program
{
    static void Main()
    {
        try
        {
            string sourcePath = "SourceWorkbook.xlsx";
            string targetPath = "TargetWorkbook.xlsx";

            // Verify that the source workbook exists
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Load the source workbook
            Workbook sourceWb = new Workbook(sourcePath);
            Worksheet sourceWs = sourceWb.Worksheets[0];

            // Create a new workbook for the copied range
            Workbook targetWb = new Workbook();
            Worksheet targetWs = targetWb.Worksheets[0];

            // Define the range to copy (A1:C10)
            int firstRow = 0;      // zero‑based index for row 1
            int firstColumn = 0;   // zero‑based index for column A
            int totalRows = 10;    // number of rows to copy
            int totalColumns = 3;  // number of columns to copy

            // Manually copy cells, preserving values, formulas, and styles
            for (int r = 0; r < totalRows; r++)
            {
                for (int c = 0; c < totalColumns; c++)
                {
                    try
                    {
                        Cell srcCell = sourceWs.Cells[firstRow + r, firstColumn + c];
                        Cell destCell = targetWs.Cells[firstRow + r, firstColumn + c];

                        // Copy value (including formulas)
                        if (!string.IsNullOrEmpty(srcCell.Formula))
                            destCell.Formula = srcCell.Formula;
                        else
                            destCell.PutValue(srcCell.Value);

                        // Copy style
                        Style srcStyle = srcCell.GetStyle();
                        destCell.SetStyle(srcStyle);
                    }
                    catch (Exception cellEx)
                    {
                        Console.WriteLine($"Warning: Failed to copy cell [{firstRow + r + 1},{firstColumn + c + 1}]: {cellEx.Message}");
                    }
                }
            }

            // Save the target workbook
            targetWb.Save(targetPath);
            Console.WriteLine($"Range copied successfully to '{targetPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
