// Title: Move a range with formulas from Sheet1 to Sheet2 and automatically refresh external references using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells move a range that contains formulas to another worksheet | How to copy cells and preserve formula references with Aspose.Cells | Recalculate workbook after moving a range between sheets in Aspose.Cells | Delete source range after copying it to a different sheet using Aspose.Cells C# | Update external formula links automatically when moving cells with Aspose.Cells
// Tags: move range with formulas Aspose.Cells | copy range between worksheets C# | recalculate formulas after range relocation | delete source range shift none Aspose.Cells | refresh external references Excel Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads or creates an Excel workbook, copies a formula‑filled range (A1:C10) from Sheet1 to D5 on Sheet2, deletes the original cells, recalculates all formulas to adjust external references, and saves the updated file.
class MoveRangeWithFormulaUpdate
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Ensure the input file exists; if not, create a minimal workbook.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
                workbook.Worksheets.Add("Sheet2");
                workbook.Save(inputPath);
            }

            // Get source and destination worksheets
            Worksheet sheet1 = workbook.Worksheets["Sheet1"];
            Worksheet sheet2 = workbook.Worksheets["Sheet2"];

            // Define the range to move (A1:C10)
            int firstRow = 0;      // zero‑based index for row 1
            int firstColumn = 0;   // zero‑based index for column A
            int totalRows = 10;    // number of rows to move
            int totalColumns = 3;  // number of columns to move (A‑C)

            // Destination start cell in Sheet2 (D5)
            int destRow = 4;       // zero‑based index for row 5
            int destColumn = 3;    // zero‑based index for column D

            // Copy the range from Sheet1 to Sheet2
            Aspose.Cells.Range srcRange = sheet1.Cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns);
            Aspose.Cells.Range destRange = sheet2.Cells.CreateRange(destRow, destColumn, totalRows, totalColumns);
            srcRange.Copy(destRange);

            // Delete the original range from Sheet1 (move operation)
            sheet1.Cells.DeleteRange(firstRow, firstColumn, totalRows, totalColumns, ShiftType.None);

            // Recalculate formulas to reflect any changes
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
